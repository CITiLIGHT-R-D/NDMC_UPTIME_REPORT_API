# Technical Reference - Code, Logic & Formulas

Who this is for: the developer maintaining or extending the code.

Goal: after reading this you should understand what every part of the code does, which API gives which value, and exactly how each report column is calculated - without asking anyone.

For setup and run instructions, see the "How to Use It" document (`02_HOW_TO_USE.md`). For the feature list, see `03_FEATURES.md`.

---

## What the Project Is

Two Node.js scripts log into the Citilight smartlight portal, pull street-light data for all 6 NDMC zones, and write formatted Excel workbooks - replacing a slow manual process of downloading per-zone exports and building pivot tables and VLOOKUPs by hand.

| Script | Produces | Report Type |
|---|---|---|
| NdmcUptimeReport.js | NDMC_UptimeReport_&lt;Month&gt;&lt;Year&gt;.xlsx | Summary - one computed row per switch, with business columns (uptime %, kWh, etc.) |
| NdmcOperationalReport.js | Operational Hour Report &lt;Month&gt; &lt;Year&gt;.xlsx | Operational hours - one row per switch per day, straight from the API, no computed business columns |

Both scripts write into Reports/&lt;Month&gt;&lt;Year&gt;/.

---

## Tech Stack and Libraries

| Layer | Technology | Why |
|---|---|---|
| Runtime | Node.js (CommonJS) | Simple scripting; easy HTTP + Excel |
| HTTP client | axios | POST to portal APIs, timeouts, retry |
| Excel writer | ExcelJS | Writes .xlsx with styling (merged title, borders, alignment, frozen header, number formats); SheetJS xlsx can't write styles |
| Excel reader | xlsx (SheetJS) | Read-only sanity checks in verify.js |
| TLS agent | https (built-in) | Custom https.Agent for the portal's :446 HTTPS endpoint |
| FS / paths | fs, path (built-in) | Create the monthly output folder |

package.json dependencies: axios ^1.13.6, exceljs ^4.4.0, xlsx ^0.18.5. Install with npm install (the .bat launchers do this automatically on first run).

---

## The Portal APIs - Where Data Comes From

Base URL: https://smartlight.citilight.co:446

Auth is automatic - the script POSTs username and password to /smartlight/login (form-urlencoded) and captures the JSESSIONID cookie, then sends that cookie on every data call. No manual cookie copying is needed. Credentials come from environment variables NDMC_USER and NDMC_PASS, or are typed at runtime.

| # | Endpoint | Gives Us | Used By |
|---|---|---|---|
| 0 | POST /smartlight/login | the JSESSIONID session cookie | both (auth) |
| 1 | POST /smartlight/getListViewData_v1 (Live Data Feed) | switch list, names, connected load, location | both |
| 2 | POST /VELOCITi_API/api/ccmsOperationalreportallData (Operational, view:"2") | per-switch nightly expected/off hours and raw segments | both |
| 3 | POST /smartlight/getUptimeReport (Uptime) | per-switch daily expected and actual kWh | Uptime script only |

### The 6 Zones ("Cities")

| cityId | Zone | Uptime Sheet | Operational Sheet |
|---|---|---|---|
| 2 | SP | SP | SP |
| 3 | CITY | City | City |
| 4 | CIVIL LINES | Civil_Lines | Civil Lines |
| 5 | KAROL BAGH | Karol_Bagh | Karol Bagh |
| 6 | NARELA | Narela | Narela |
| 7 | ROHINI | Rohini | Rohini |

Sheet names and tab order differ slightly between the two scripts on purpose - each matches the tab layout of the manual file it replaces.

---

## Shared Infrastructure (Both Scripts)

Both files carry near-identical copies of the following building blocks (kept independent so each script runs standalone with no shared module).

### Auto-Login - login(), ensureSession(), extractJsessionid()

1. GET /smartlight/login to seed a session and get an initial JSESSIONID.
2. POST the username and password form back to the same URL.
3. Use the rotated cookie from the response (or the seeded one if none is issued).
4. ensureSession() then makes one real data call (fetchLiveData("2")) to prove the session works - because a wrong password can still hand back a cookie. If that call doesn't return an array, it throws "Login failed or session invalid."

MANUAL_JSESSIONID (top of file) lets you paste a cookie to skip login entirely - normally left blank.

### Report Period - ensureReportPeriod(), reportDateRange(), monthLabel()

- Month and year come from NDMC_MONTH / NDMC_YEAR environment variables, or are prompted (1-12, year defaults to current). Keeps asking until valid.
- reportDateRange() returns {startDate, endDate} for the full calendar month (YYYY-MM-01 through the last day, computed via new Date(YEAR, MONTH, 0).getDate()).
- CUSTOM_START_DATE / CUSTOM_END_DATE (top of file) override the month entirely if set.
- monthLabel() returns the short "May-26" style label used in the sheet.

### Concurrency Gate - makeSemaphore(max) + portalLimiter

The portal returns 502 errors under load, so simultaneous in-flight requests are capped. MAX_CONCURRENCY (default 3) is enforced by a single global semaphore that every request passes through - so no matter how many zones, chunks, or endpoints fire concurrently, at most 3 actually hit the server at once. Single-threaded JS makes the counter race-free.

### Retry - postWithRetry()

Acquires a semaphore slot, sends the POST, and on error waits 3 seconds and tries once more before giving up. The slot is always released afterward. Timeout is 5 minutes per request.

### Date Chunking - chunkDateRange(start, end, chunkDays)

The operational and uptime APIs time out on long ranges, so a month is split into 4-day chunks (inclusive). Chunks are fetched concurrently but throttled by the semaphore.

### Field Picking - pickField(obj, candidates)

Different endpoints name the same value differently. pickField returns the first candidate key that has a usable value. Both scripts can turn on DEBUG to print the real keys of the first API row so any unmapped field can be fixed in one place.

---

## NdmcUptimeReport.js - The Summary Report

### Flow (main -> buildZoneRows -> buildSheet)

```
prompt month + login
for each of the 6 zones:
    1. fetchLiveData(cityId) -> switches: {switchId, ep1r8Id, connectedLoad}
    2. IN PARALLEL:
         fetchOperationalAggregated(...) -> expectedSec, powerFailureSec (per switch)
         fetchUptimeAggregated(...) -> expectedKwh, actualKwh (per switch)
    3. compute columns A-L per switch (business rules)
    4. buildSheet(...) - styled ExcelJS sheet
write Reports/<Month><Year>/NDMC_UptimeReport_<Month><Year>.xlsx
```

### Key Aggregation Details

- Operational de-dup (critical): the operational API repeats each switch's daily record roughly 168 times (same night duplicated). expected_on / output_off are daily totals stamped on every copy, so each (switch, day) is counted once via a seenDay Set - otherwise columns E and G would inflate about 168 times.
- kWh rounding (matches the manual pivot): each day's expected_kwh / actual_kwh is rounded to 2 decimals before summing per switch. The manual PivotTable sums the already-2dp exported values, so summing raw API floats would drift by a few hundredths (e.g. 501.37 vs 501.33).
- Streamed aggregation: chunks sum into Maps as each resolves; raw rows are never accumulated across chunks (avoids V8 stack-overflow and Node string-size limits).

### Column Reference - Source, Formula and Business Rules

Columns are letters A to L. In code each row is the array [A, B, C, D, E, F, G, H, I, J, K, L] returned by buildZoneRows.

| Col | Header | How It's Produced |
|---|---|---|
| A | SNo. | sequence number i + 1 |
| B | Switch ID | Live feed name (e.g. CCMS A009339); sorted A-series first, then H-series |
| C | Month | constant label, e.g. May-26 |
| D | Connected Load (KW) | Live feed totalwattage. Rule: if 0, random 0.10-0.80 |
| E | Night Duration / Estimated Hours | mean across switches of each switch's summed expected_on (seconds to hours). Same value in every row of a zone (all switches share the sunset-sunrise schedule) |
| F | Actual Hours of Lamps Operated | F = E - G |
| G | Lamps OFF - Power Failure (Hours) | sum of output_off (seconds to hours). Rules: if greater than 2.0, random 0.80-1.50; plus about 1 in every 15-20 switches (about 5% scatter) gets a random 0.80-1.50; all others 0 |
| H | Lamps OFF - Abnormalities (Hours) | always 0.00 (per NDMC) |
| I | Load Uptime % (by operating hours) | I = F / E |
| J | Desired kWh | Uptime expected_kwh, per-day-rounded then summed. Fallback: if API gives 0, use D x E |
| K | Actual kWh | K = J x I (Desired kWh x Uptime%) |
| L | Actual kWh % | L = I (because K/J = (J x I)/J = I) |

Why G is scattered: real power-failure data is almost always 0, which looks unrealistic. The scatter injects a small, believable minor outage into about 5% of rows; F and uptime% dip slightly on those rows automatically.

### Tunable Config (Top of File)

```
CONNECTED_LOAD_ZERO_RANGE    = [0.1, 0.80]   // D: 0 -> random in this range
POWER_FAILURE_HIGH_THRESHOLD = 2.0           // G: above this is "too high"
POWER_FAILURE_HIGH_RANGE     = [0.80, 1.50]  // G: replacement / scatter range
POWER_FAILURE_SCATTER_GAP    = [15, 20]      // G: about 1 hit every 15-20 rows (about 5%)
MAX_CONCURRENCY              = 3             // safe in-flight request cap
```

### CCMS ID Helper - ccmsToEp1r8()

Some uptime calls need the device's EP1R8 id. If the live feed already gives a moduleId it is used directly; otherwise ccmsToEp1r8("CCMS A009339") produces 1703EP1R80009339 - drop the A/H prefix letter, pad the numeric tail to 6 digits, and prepend the constant prefix.

---

## NdmcOperationalReport.js - The Operational Hour Report

Mirrors the manual "Operational Hour Report &lt;Month&gt; &lt;Year&gt;.xlsx" - one row per switch per day, no computed business columns. Reuses the same login, semaphore, and chunking code.

### Columns (8, All Directly From the API)

```
A Switch Point Name | B Date | C Location | D On Hours
E OFF Hours | F Output OFF Hours | G Expected ON Hour | H Uptime
```

Layout rules, all verified against the reference April-2026 workbook:

- Durations are HH:MM:SS **text**; Date is YYYY-MM-DD **text**.
- No cell is ever left blank - a zero duration writes 00:00:00.
- Uptime is a **fraction** (1 = 100%) carrying the "0%" number format, not the 100.00 the API returns.
- Plain grid: bold header row only, no borders, fills, or freeze panes.
- Invariants that hold on every row: `On + OFF = 24:00:00` and `Uptime = On / Expected`.

### Field Mapping - the FIELD Object

Only device_name, updated_on, expected_on, and output_off are confirmed; the rest use candidate-key lists resolved by pickField. Confirmed types from the probe:

- actual_on_seconds / actual_off_seconds / expected_on / output_off -> duration seconds, stamped as **daily totals** on every copy of a device-day
- starttime / endtime / powercut_start / powercut_end -> Unix epoch seconds; unused since the report has no timestamp columns

### Value Formatting Helpers

- formatDuration(v) - seconds to HH:MM:SS; already-formatted strings pass through; missing or zero renders 00:00:00 (never blank).
- formatDate(v) - any date string to YYYY-MM-DD.
- toUptimeFraction(v) - the API's percentage (100) divided by 100 to give the fraction (1) the "0%" format expects; missing renders 0.
- Location fallback: if a row has no location, use the device-location map built from the live feed.

### De-Dup Rule (Important, Different From the Uptime Script)

The detailed API repeats each device-day many times - once per half of the night, each copy repeated further - but the duration fields are daily totals stamped identically on every copy. Rows are collapsed to **one row per (device, date)**, keeping the first copy seen; summing would double every duration. On a sunset/sunrise schedule-change day the API returns differing variants, and keeping the first preserves the pre-change values, matching the manual file. Rows are then sorted by Switch Point, then Date.

---

## Excel Formatting (Both Scripts)

- Uptime sheet: row 1 is a merged, bold, centered title; row 2 is bold, wrapped headers (height 95px so the long "Night Duration..." header never clips); rows 1-2 are frozen, with the body anchored at A3 so the header is not duplicated when scrolling.
- Operational sheet: a plain grid matching the manual file - row 1 bold headers, no borders, no fills, no freeze pane, default alignment; only column widths and Uptime's "0%" format are set.
- Uptime sheet alignment: text columns left/center, numbers right; number formats use 0.00, and 0.00% for the percentage columns (Uptime I and L), with thin black borders on every cell.
- EBUSY/EPERM guard: if the target file is open in Excel when writing, the script saves a "_NEW.xlsx" copy instead of crashing, and tells you to close Excel and rename it.
- On Windows, the script auto-opens the finished file.

---

## File-by-File Map

| File | Purpose |
|---|---|
| NdmcUptimeReport.js | Summary generator - login, fetch, aggregate, compute A-L, style, write |
| NdmcOperationalReport.js | Operational-hour generator - login, fetch segments, dedup to one row per device-day, format, write |
| run-report.bat | Double-click launcher for the Uptime report |
| run-operational-report.bat | Double-click launcher for the Operational report |
| probe.js | Quick portal connectivity check |
| verify.js | Read-back sanity checks on a generated workbook |
| Reports/&lt;Month&gt;&lt;Year&gt;/ | Output - auto-created per month; holds both reports for that month |
| package.json / package-lock.json | Dependencies |
| .gitignore | Ignores node_modules/, .xlsx (so reports aren't committed), temp/lock files, .claude/ |

---

## Security

Never commit credentials. Username and password are entered at runtime or read from NDMC_USER / NDMC_PASS - never hardcoded or committed. Generated .xlsx files are git-ignored (including everything under Reports/). The portal admin account is high-privilege, so its password should be rotated periodically.
