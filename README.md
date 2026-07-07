# NDMC_UPTIME_REPORT_API

Automates the **NDMC Monthly Uptime & Energy Consumption Report** — generates
`NDMC_UptimeReport_<Month><Year>.xlsx` with one sheet per zone (SP, City, Civil_Lines,
Karol_Bagh, Narela, Rohini), pulled directly from the Citilight smartlight portal APIs.

## What it does

In a single run it fetches, for all 6 zones:
- **Live Data Feed** → Switch ID + Connected Load (columns B, D)
- **Operational Report** → Night/expected hours, Power Failure (columns E, G)
- **Uptime Report** → Desired & Actual kWh (columns J, K)

It then computes the formula columns (F = E−G, I = F/E, L = K/J), applies the
business rules, and writes the finished, fully-styled workbook (via **ExcelJS**) —
no manual pivots or VLOOKUPs.

## 📘 Documentation (start here for KT)

Three focused guides in [`docs/`](docs/) — read them in order:

| Doc | For | Covers |
|---|---|---|
| [`docs/01_TECHNICAL.md`](docs/01_TECHNICAL.md) | developers | how the code works, every API source, all column formulas & business rules |
| [`docs/02_HOW_TO_USE.md`](docs/02_HOW_TO_USE.md) | anyone running it | step-by-step run guide, prompts, verify checklist, troubleshooting |
| [`docs/03_FEATURES.md`](docs/03_FEATURES.md) | overview | full feature list (incl. the new monthly output folders) |

> The older [`PROJECT_OVERVIEW.md`](PROJECT_OVERVIEW.md) and [`RUN_GUIDE.md`](RUN_GUIDE.md)
> remain as background references; the `docs/` set above is the current, complete source.

## Where reports are saved

Each run auto-creates `Reports\<Month><Year>\` and writes the report there — e.g.
`Reports\Jun2026\NDMC_UptimeReport_Jun2026.xlsx`. Both reports for a month share one folder.

## Usage

**No cookie needed** — the tool logs in for you. Just run it and answer 3 prompts.

- **Easiest:** double-click **`run-report.bat`**.
- **Or PowerShell:**
  ```powershell
  cd D:\CityReport
  node NdmcUptimeReport.js
  ```

It asks: **which month**, which year, and your portal **username / password** — then logs
in automatically, generates the report (~10–15 min), and opens the Excel file. To skip the
prompts (automation), set `NDMC_MONTH`, `NDMC_YEAR`, `NDMC_USER`, `NDMC_PASS` env vars.

Full step-by-step instructions, troubleshooting, and the column source mapping are in
[`RUN_GUIDE.md`](RUN_GUIDE.md).

## Performance

Requests are fetched concurrently (throttled by a global semaphore, default 3) with
operational + uptime running in parallel per zone. A full month runs in ~10–15 min.

## Security

⚠️ **Never commit credentials.** The portal username/password are entered at runtime (or
read from `NDMC_USER` / `NDMC_PASS` env vars) — they are never hardcoded in the code or
committed. Generated `.xlsx` files are git-ignored.

## Files

| File | Purpose |
|---|---|
| `NdmcUptimeReport.js` | Main report generator (auto-login + styled Excel output) |
| `run-report.bat` | Double-click launcher: prompts for month + credentials, runs, opens the file |
| `probe.js` | Quick portal connectivity check |
| `verify.js` | Sanity-check the generated workbook's columns |
| `PROJECT_OVERVIEW.md` | Complete reference: tech, libraries, flow, columns, commands |
| `RUN_GUIDE.md` | Full monthly run instructions |
