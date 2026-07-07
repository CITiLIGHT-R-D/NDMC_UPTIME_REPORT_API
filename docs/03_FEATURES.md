# 03 — Features

> A plain list of what this tool does and the behaviours built into it — so during a KT
> handover nobody has to reverse-engineer the code to learn "what can it do?".

See also: [`01_TECHNICAL.md`](01_TECHNICAL.md) (how) · [`02_HOW_TO_USE.md`](02_HOW_TO_USE.md) (run it).

---

## Core features

| # | Feature | What it means |
|---|---|---|
| 1 | **Two reports** | Uptime (summary, 1 row/switch, computed columns) and Operational (detailed raw nightly segments). |
| 2 | **All 6 zones in one run** | SP, City, Civil Lines, Karol Bagh, Narela, Rohini — one sheet each, one workbook. |
| 3 | **Auto-login (no cookies)** | Logs into the portal with your username/password and grabs the `JSESSIONID` itself. No copying cookies from the browser. |
| 4 | **Interactive month prompt** | Just asks "which month / year". Or preset with env vars for automation. |
| 5 | **Fully styled Excel output** | Merged title, bold wrapped headers, frozen header row, borders, 2-decimal & percentage number formats — matches the manual NDMC layout. |
| 6 | **Automatic business rules & formulas** | Uptime %, kWh, power-failure substitution, etc. all computed for you (see [`01_TECHNICAL.md` §5.3](01_TECHNICAL.md)). No manual pivots/VLOOKUPs. |
| 7 | **One-click launchers** | `run-report.bat` and `run-operational-report.bat` — double-click, answer 3 questions, done. Installs dependencies on first run. |
| 8 | **Auto-opens the finished file** | On Windows the report opens in Excel when generation finishes. |

## ⭐ New: automatic monthly output folders

**What it does:** every time you generate a report it now creates (if needed) a folder
`Reports\<Month><Year>\` and saves the report inside it — instead of leaving files loose in
the project root.

- Example: a June 2026 run writes `Reports\Jun2026\NDMC_UptimeReport_Jun2026.xlsx`.
- Both the Uptime and the Operational report for the same month land in the **same** month
  folder, so everything for a period is together.
- The folder is created automatically (`fs.mkdirSync(..., {recursive:true})`) — no manual
  setup, and re-running the same month reuses the existing folder without error.
- Existing older reports were moved into this structure so the folders are aligned:
  ```
  Reports/
    May2026/
      NDMC_UptimeReport_May2026.xlsx
      Operation_uptime_Reports_May2026.xlsx
    Jun2026/
      NDMC_UptimeReport_Jun2026.xlsx
      Operational Hour Report June 2026.xlsx
    <NextMonth>/           ← created automatically on the next run
  ```

**Why it was added:** previously all `.xlsx` files piled up in the project root and were hard
to tell apart across months. Now each month is self-contained and easy to find/archive.

---

## Reliability features (built in, nothing to configure)

| Feature | What it protects against |
|---|---|
| **Concurrency cap (semaphore)** | The portal 502s under load — at most `MAX_CONCURRENCY` (default 3) requests hit it at once, across all zones/chunks. |
| **Date chunking (4-day windows)** | The APIs time out on long ranges — a month is fetched in small concurrent chunks. |
| **Retry on failure** | Each request retries once after a 3 s pause before giving up. |
| **Session verification** | After login it makes a real data call to confirm the session actually works (a wrong password can still return a cookie). |
| **De-duplication** | The APIs repeat records many times; the code counts each switch-day / night-segment once so totals aren't inflated. |
| **kWh rounding to match the manual pivot** | Each day is rounded to 2dp before summing, so totals match the old hand-built PivotTable exactly. |
| **"File open in Excel" guard** | If the target file is open when writing, it saves a `_NEW.xlsx` copy instead of crashing. |
| **Per-zone isolation** | If one zone fails, the others still complete and the run finishes with a summary. |

---

## Configurable knobs (top of each script)

| Setting | Default | Effect |
|---|---|---|
| `MAX_CONCURRENCY` | `3` | Speed vs. portal safety. Raise to 5 only if runs show 0 failed chunks. |
| `CONNECTED_LOAD_ZERO_RANGE` | `0.10–0.80` | Random value used when Connected Load is 0. |
| `POWER_FAILURE_HIGH_THRESHOLD` | `2.0` | Above this, power-failure hours are treated as unrealistic and replaced. |
| `POWER_FAILURE_HIGH_RANGE` | `0.80–1.50` | Replacement / scatter range for power-failure hours. |
| `POWER_FAILURE_SCATTER_GAP` | `15–20` | ~1 in every 15–20 switches gets a small scattered outage (≈5%). |
| `CUSTOM_START_DATE` / `CUSTOM_END_DATE` | `""` | Set both to report an explicit date range instead of a full month. |
| `MANUAL_JSESSIONID` | `""` | Paste a cookie to skip auto-login (rarely needed). |

Env-var overrides (skip the prompts): `NDMC_MONTH`, `NDMC_YEAR`, `NDMC_USER`, `NDMC_PASS`.

---

## Not included / by design

- **Credentials are never stored** — typed at runtime or supplied via env vars, never in code.
- **Reports are not committed to git** — `*.xlsx` (including everything under `Reports/`) is
  git-ignored; the repo holds code + docs only.
- **Column H (Abnormalities) is always 0.00** — per NDMC reporting rules.
