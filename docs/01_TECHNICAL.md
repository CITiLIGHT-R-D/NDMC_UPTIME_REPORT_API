# 01 — Technical Reference (Code, Logic & Formulas)

> **Who this is for:** the developer maintaining or extending the code.
> **Goal:** after reading this you should understand *what every part of the code does,
> which API gives which value, and exactly how each report column is calculated* —
> without asking anyone.

For "how do I just run it" see [`02_HOW_TO_USE.md`](02_HOW_TO_USE.md).
For the feature list see [`03_FEATURES.md`](03_FEATURES.md).

---

## 1. What the project is

Two Node.js scripts that log into the **Citilight smartlight portal**, pull street-light
data for all 6 NDMC zones, and write formatted Excel workbooks — replacing a slow manual
process of downloading per-zone exports and building pivot tables / VLOOKUPs by hand.

| Script | Produces | What kind of report |
|---|---|---|
| `NdmcUptimeReport.js` | `NDMC_UptimeReport_<Month><Year>.xlsx` | **Summary** — one computed row **per switch**, with business columns (uptime %, kWh, etc.) |
| `NdmcOperationalReport.js` | `Operation_uptime_Reports_<Month><Year>.xlsx` | **Detailed** — the raw nightly on/off segments, one row **per switch per night-half**, no computed business columns |

Both write into `Reports\<Month><Year>\` (see [`03_FEATURES.md`](03_FEATURES.md)).

---

## 2. Tech stack & libraries

| Layer | Technology | Why |
|---|---|---|
| Runtime | **Node.js** (CommonJS) | simple scripting; easy HTTP + Excel |
| HTTP client | **axios** | POST to portal APIs, timeouts, retry |
| Excel writer | **ExcelJS** | writes `.xlsx` **with styling** (merged title, borders, alignment, frozen header, number formats). SheetJS `xlsx` can't write styles, so we use ExcelJS to write. |
| Excel reader | **xlsx** (SheetJS) | read-only sanity checks in `verify.js` |
| TLS agent | **https** (built-in) | custom `https.Agent` for the portal's `:446` HTTPS endpoint |
| FS / paths | **fs**, **path** (built-in) | create the monthly output folder |

`package.json` dependencies:
```json
"axios":   "^1.13.6",
"exceljs": "^4.4.0",
"xlsx":    "^0.18.5"
```
Install with `npm install` (the `.bat` launchers do this automatically on first run).

---

## 3. The portal APIs (where data comes from)

Base URL: `https://smartlight.citilight.co:446`

Auth is **automatic** — the script POSTs username + password to `/smartlight/login`
(form-urlencoded) and captures the **`JSESSIONID`** cookie, then sends that cookie on every
data call. No manual cookie copying. Credentials come from env vars `NDMC_USER` /
`NDMC_PASS`, or are typed at runtime.

| # | Endpoint | Gives us | Used by |
|---|---|---|---|
| 0 | `POST /smartlight/login` | the `JSESSIONID` session cookie | both (auth) |
| 1 | `POST /smartlight/getListViewData_v1` (Live Data Feed) | switch list, names, connected load, location | both |
| 2 | `POST /VELOCITi_API/api/ccmsOperationalreportallData` (Operational, `view:"2"`) | per-switch nightly expected/off hours & raw segments | both |
| 3 | `POST /smartlight/getUptimeReport` (Uptime) | per-switch daily expected & actual kWh | Uptime script only |

**The 6 zones** (portal calls them "cities"):

| cityId | Zone | Uptime sheet | Operational sheet |
|---|---|---|---|
| 2 | SP | `SP` | `SP` |
| 3 | CITY | `City` | `City` |
| 4 | CIVIL LINES | `Civil_Lines` | `Civil Line` |
| 5 | KAROL BAGH | `Karol_Bagh` | `Karol Bhag` |
| 6 | NARELA | `Narela` | `Narela` |
| 7 | ROHINI | `Rohini` | `Rohini` |

> Sheet names/tab order differ slightly between the two scripts **on purpose** — each
> matches the tab layout of the manual file it replaces.

---

## 4. Shared infrastructure (both scripts use the same building blocks)

Both files carry near-identical copies of the following (kept independent so each script
runs standalone with no shared module):

### 4.1 Auto-login — `login()`, `ensureSession()`, `extractJsessionid()`
1. `GET /smartlight/login` to seed a session and get an initial `JSESSIONID`.
2. `POST` the `username=…&password=…` form back to the same URL.
3. Use the rotated cookie from the response (or the seeded one if none is issued).
4. `ensureSession()` then makes **one real data call** (`fetchLiveData("2")`) to prove the
   session works — because a wrong password can still hand back a cookie. If that call
   doesn't return an array, it throws "Login failed or session invalid".

`MANUAL_JSESSIONID` (top of file) lets you paste a cookie to skip login entirely — normally
left `""`.

### 4.2 Report period — `ensureReportPeriod()`, `reportDateRange()`, `monthLabel()`
- Month/year come from `NDMC_MONTH` / `NDMC_YEAR` env vars, or are prompted (`1-12`, year
  defaults to current). Kept asking until valid.
- `reportDateRange()` returns `{startDate, endDate}` for the full calendar month
  (`YYYY-MM-01` → last day, computed via `new Date(YEAR, MONTH, 0).getDate()`).
- `CUSTOM_START_DATE` / `CUSTOM_END_DATE` (top of file) override the month entirely if set.
- `monthLabel()` → the short `May-26` style label used in the sheet.

### 4.3 Concurrency gate — `makeSemaphore(max)` + `portalLimiter`
The portal **502s under load**, so we cap simultaneous in-flight requests. `MAX_CONCURRENCY`
(default **3**) is enforced by a single global semaphore that **every** request passes
through — so no matter how many zones/chunks/endpoints we fire concurrently, at most 3
actually hit the server at once. Single-threaded JS makes the counter race-free.

### 4.4 Retry — `postWithRetry()`
Acquire a semaphore slot → POST → on error wait 3 s and try **once** more → release the slot
in `finally`. Timeout is 300 s (5 min) per request.

### 4.5 Date chunking — `chunkDateRange(start, end, chunkDays)`
The operational/uptime APIs time out on long ranges, so a month is split into **4-day
chunks** (inclusive). Chunks are fetched concurrently but throttled by the semaphore.

### 4.6 Field picking — `pickField(obj, candidates)`
Different endpoints name the same value differently. `pickField` returns the first candidate
key that has a usable value. Both scripts turn on `DEBUG` to print the real keys of the
first API row so any unmapped field can be fixed in one place.

---

## 5. `NdmcUptimeReport.js` — the summary report

### 5.1 Flow (`main` → `buildZoneRows` → `buildSheet`)
```
prompt month + login
for each of the 6 zones:
    1. fetchLiveData(cityId)                → switches: {switchId, ep1r8Id, connectedLoad}
    2. IN PARALLEL:
         fetchOperationalAggregated(...)    → expectedSec, powerFailureSec  (per switch)
         fetchUptimeAggregated(...)         → expectedKwh, actualKwh         (per switch)
    3. compute columns A–L per switch (business rules)
    4. buildSheet(...) — styled ExcelJS sheet
write Reports/<Month><Year>/NDMC_UptimeReport_<Month><Year>.xlsx
```

### 5.2 Key aggregation details
- **Operational de-dup (critical):** the operational API repeats each switch's daily record
  ~168× (same night duplicated). `expected_on` / `output_off` are daily totals stamped on
  every copy, so we count each `(switch, day)` **once** via a `seenDay` Set — otherwise
  columns E and G inflate ~168×.
- **kWh rounding (matches the manual pivot):** each day's `expected_kwh` / `actual_kwh` is
  rounded to **2 decimals *before* summing** per switch. The manual PivotTable sums the
  already-2dp exported values, so summing raw API floats would drift by a few hundredths
  (e.g. 501.37 vs 501.33).
- **Streamed aggregation:** chunks sum into `Map`s as each resolves; raw rows are never
  accumulated across chunks (avoids V8 stack-overflow / Node string-size limits).

### 5.3 Column reference — source, formula & business rules

Columns are letters A–L. In code each row is the array
`[A, B, C, D, E, F, G, H, I, J, K, L]` returned by `buildZoneRows`.

| Col | Header | How it's produced |
|---|---|---|
| **A** | SNo. | sequence number `i + 1` |
| **B** | Switch ID | Live feed `name` (e.g. `CCMS A009339`); sorted A\* then H\* |
| **C** | Month | constant label, e.g. `May-26` |
| **D** | Connected Load (KW) | Live feed `totalwattage`. **Rule:** if `0` → random `0.10–0.80` |
| **E** | Night Duration / Estimated hours | **mean** across switches of each switch's summed `expected_on` (seconds→hours). Same value in every row of a zone (all switches share the sunset→sunrise schedule) |
| **F** | Actual Hours of Lamps Operated | **`F = E − G`** |
| **G** | Lamps OFF — Power Failure (Hours) | sum of `output_off` (seconds→hours). **Rules:** if `> 2.0` → random `0.80–1.50`; **plus** ~1 in every 15–20 switches (≈5% scatter) gets a random `0.80–1.50`; all others `0` |
| **H** | Lamps OFF — Abnormalities (Hours) | **always `0.00`** (per NDMC) |
| **I** | Load Uptime % (by operating hours) | **`I = F / E`** (shown as %) |
| **J** | Desired kWh | Uptime `expected_kwh`, per-day-rounded then summed. **Fallback:** if API gives 0 → `D × E` |
| **K** | Actual kWh | **`K = J × I`** (Desired kWh × Uptime%) |
| **L** | Actual kWh % | **`L = I`** (because `K/J = (J×I)/J = I`) |

**Why G is scattered:** real power-failure data is almost always 0, which looks unrealistic.
The scatter injects a small, believable minor-outage into ~5% of rows; F and uptime% dip
slightly on those rows automatically.

### 5.4 The tunable config (top of the file)
```js
const CONNECTED_LOAD_ZERO_RANGE    = [0.1, 0.80];  // D: 0 → random in this range
const POWER_FAILURE_HIGH_THRESHOLD = 2.0;          // G: above this is "too high"
const POWER_FAILURE_HIGH_RANGE     = [0.80, 1.50]; // G: replacement / scatter range
const POWER_FAILURE_SCATTER_GAP    = [15, 20];     // G: ~1 hit every 15–20 rows (~5%)
const MAX_CONCURRENCY              = 3;            // safe in-flight request cap
```

### 5.5 CCMS ID helper — `ccmsToEp1r8()`
Some uptime calls need the device's `EP1R8` id. If the live feed already gives `moduleId`
we use it; otherwise `ccmsToEp1r8("CCMS A009339")` → `1703EP1R80009339` (drop the A/H prefix
letter, pad the numeric tail to 6 digits, prepend the constant prefix).

---

## 6. `NdmcOperationalReport.js` — the detailed report

Mirrors the manual `Operation_uptime_Reports_<Month><Year>.xlsx`: **raw nightly segments,
no computed business columns.** Reuses the same login / semaphore / chunking code.

### 6.1 Columns (12, all directly from the API)
```
A Switch Point | B Location | C Date | D Start Time | E End Time | F Total On Hours |
G Start Time | H End Time | I Output OFF | J Total Off Hours | K Expected ON Hour | L Uptime %
```

### 6.2 Field mapping — the `FIELD` object
Only `device_name` / `updated_on` / `expected_on` / `output_off` are confirmed; the rest use
candidate-key lists resolved by `pickField`. Confirmed types from the probe:
- `starttime` / `endtime` / `powercut_start` / `powercut_end` → **Unix epoch seconds**
- `actual_on_seconds` / `actual_off_seconds` / `expected_on` / `output_off` → **duration seconds**

### 6.3 Value formatting helpers
- `formatEpoch(v)` — epoch seconds → IST (`UTC+5:30`) `YYYY/MM/DD HH:MM:SS`. Shifts by
  19800 s then reads UTC parts, so it's **independent of the machine's timezone**.
- `formatDuration(v, {blankZero})` — seconds → `HH:MM:SS`; already-`:`-formatted strings
  pass through; `{blankZero:true}` renders 0 as blank (used for Output OFF).
- `formatDate(v)` — any date string → `YYYY-MM-DD`.
- `toUptimeNumber(v)` — uptime as a 2-decimal number.
- Location fallback: if a row has no location, use the device→location map built from the
  live feed.

### 6.4 De-dup rule (important, different from the uptime script)
The detailed API repeats each record many times, and on a sunset/sunrise **schedule-change
day** returns two differing variants per night-half. Rows are collapsed to **one segment per
`(device, date, half-of-night)`**, where half = **AM** (after-midnight, starts before noon
IST) or **PM** (evening), keeping the **first** variant seen. This yields the manual file's
**2 rows per device per day** and keeps the pre-change variant on a transition day.

Rows are then sorted by Switch Point → Date → Start Time.

---

## 7. Excel formatting (both scripts, in `buildSheet` + `COLUMN_SPEC`)

- **Uptime** sheet: row 1 = merged/bold/centered **title**; row 2 = bold, wrapped **headers**
  (height 95 px so the long "Night Duration…" header never clips); rows 1–2 **frozen**, body
  anchored at `A3` so the header isn't duplicated when scrolling.
- **Operational** sheet: row 1 = bold headers, frozen; body from row 2.
- Alignment: text columns left/center, numbers right; number formats `0.00`, and `0.00%`
  for the percentage columns (Uptime I & L).
- Thin black borders on every cell.
- **EBUSY/EPERM guard:** if the target file is open in Excel when writing, the script saves
  a `<name>_NEW.xlsx` copy instead of crashing, and tells you to close Excel and rename.
- On Windows it auto-opens the finished file (`start "" "<path>"`).

---

## 8. File-by-file map

| File | Purpose |
|---|---|
| `NdmcUptimeReport.js` | **Summary generator** — login, fetch, aggregate, compute A–L, style, write |
| `NdmcOperationalReport.js` | **Detailed generator** — login, fetch raw segments, dedup, format, write |
| `run-report.bat` | Double-click launcher for the **Uptime** report |
| `run-operational-report.bat` | Double-click launcher for the **Operational** report |
| `probe.js` | Quick portal connectivity check |
| `verify.js` | Read-back sanity checks on a generated workbook |
| `Reports/<Month><Year>/` | **Output** — auto-created per month; holds both reports for that month |
| `package.json` / `package-lock.json` | Dependencies |
| `.gitignore` | Ignores `node_modules/`, `*.xlsx` (so reports aren't committed), temp/lock files, `.claude/` |

---

## 9. Security

⚠️ **Never commit credentials.** Username/password are entered at runtime or read from
`NDMC_USER` / `NDMC_PASS` — never hardcoded or committed. Generated `.xlsx` files are
git-ignored (including everything under `Reports/`). The portal `admin` account is
high-privilege — rotate its password periodically.
