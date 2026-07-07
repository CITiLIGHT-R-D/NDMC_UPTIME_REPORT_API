# 02 — How To Use (Step-by-Step Run Guide)

> **Who this is for:** anyone who just needs to *produce the monthly reports* — no coding
> required. If a new person is taking over, this is the only page they need to run it.

For how the code works internally see [`01_TECHNICAL.md`](01_TECHNICAL.md).
For the feature list see [`03_FEATURES.md`](03_FEATURES.md).

---

## What you get

Two reports, one per script:

| Report | File it creates | Saved into |
|---|---|---|
| **Uptime** (summary, 1 row per switch) | `NDMC_UptimeReport_<Month><Year>.xlsx` | `Reports\<Month><Year>\` |
| **Operational** (detailed nightly segments) | `Operation_uptime_Reports_<Month><Year>.xlsx` | `Reports\<Month><Year>\` |

Each has **6 sheets** — one per NDMC zone (SP, City, Civil Lines, Karol Bagh, Narela, Rohini).

> **Note:** the folder `Reports\<Month><Year>\` is **created automatically** when you run.
> Both reports for the same month go into the same folder. Nothing to set up by hand.

---

## One-time setup (first run only)

1. Install **Node.js** — download from https://nodejs.org, run the installer, click through.
   Check it worked: open PowerShell and type `node --version` (should print a version).
2. You need a **portal login** (username + password) for
   `https://smartlight.citilight.co:446`.
3. The launcher installs the code's dependencies for you the first time (`npm install`).
   You don't need to do anything extra.

---

## The easy way — double-click

**To make the Uptime (summary) report:**
1. Open the project folder `D:\NDMC UPTIME REPORT API`.
2. **Double-click `run-report.bat`.**
3. A black window opens and asks 3 things — type each and press **Enter**:
   ```
   Which month do you want the report for? Enter 1-12 (1=Jan … 12=Dec): 5
   Which year? Press Enter for 2026:
   Portal username: admin
   Portal password: ********        (hidden as you type)
   ```
4. It logs in, fetches all 6 zones (**~10–15 min**), and **opens the Excel automatically**
   when done. The window stays open so you can read the summary — close it when finished.

**To make the Operational (detailed) report:**
- Same thing, but **double-click `run-operational-report.bat`** instead. Same 3 questions.

That's it. Find your file in `Reports\<Month><Year>\`.

---

## The manual way — PowerShell (optional)

If you prefer the command line:

```powershell
cd "D:\NDMC UPTIME REPORT API"
node NdmcUptimeReport.js          # summary report
# or
node NdmcOperationalReport.js     # detailed report
```

It asks the same 3 questions. To **skip the prompts** (e.g. for automation/scheduling), set
environment variables first:

```powershell
$env:NDMC_MONTH = 5
$env:NDMC_YEAR  = 2026
$env:NDMC_USER  = "admin"
$env:NDMC_PASS  = "your-password"
node NdmcUptimeReport.js
```

---

## What you'll see while it runs

Live progress prints one zone at a time:
```
Report: May-26   (2026-05-01 → 2026-05-31)
Logging in…
Login OK — session acquired.

[SP] cityId=2
  live data: 133 switches
    op chunk 2026-05-01 → 2026-05-04
    ...
  uptime: 1995 daily rows (0 failed chunks)
```
At the end look for the `--- Summary ---` block with all six zones showing `OK`, and:
```
Written: Reports\May2026\NDMC_UptimeReport_May2026.xlsx
```

**While it runs:** don't close the window, and don't open the output file until it finishes.

---

## Approximate run times

| Date range | Approx time |
|---|---|
| 4 days (test) | 2–3 min |
| 15 days | 5–8 min |
| Full month | 10–15 min |

**Speed knob:** `MAX_CONCURRENCY` near the top of each script (default `3`, the safe level
for this portal). If a full run finishes with **0 failed chunks**, you may try `5`. If you
see `failed chunks` or `FAIL: 502`, put it back to `3`.

---

## Verify the result (quick checklist)

Open the output Excel. On each sheet (or just SP — the rest use the same logic) confirm:

**Uptime report:**
- Row 1: title `Monthly uptime and Energy Consumption Report <Zone> Zone - <Month><Year>`.
- Row 2: 12 headers (A SNo. … L Actual kWh %), frozen so they stay visible.
- Column B: switch IDs like `CCMS A009339`, never blank; sorted A\* then H\*.
- Column C: the month label in every row (e.g. `May-26`).
- Column D: values > 0 (zeros auto-replaced with 0.10–0.80).
- Column E: the **same** value in every row of the sheet.
- Column F = E − G; Column H always `0.00`; Column I = F/E as %; Column L = K/J as %.
- Column G: mostly 0; ~5% of rows have a small value, all ≤ ~1.50.

Switch counts per zone (May 2026 baseline, ±2):
SP 133, City 143, Civil_Lines 823, Karol_Bagh 341, Narela 884, Rohini 962.

You can also run `node verify.js` after a run for automated row-count / rule checks.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `Login failed or session invalid` | Wrong username/password | Re-run, re-type carefully |
| `login: server did not return a JSESSIONID cookie` | Portal changed/unreachable | Confirm the portal loads in a browser |
| Zone says `FAIL: 502` | Portal overloaded | Wait 5 min, retry. Persistent = portal down |
| Zone says `FAIL: timeout` | A call took > 5 min | Retry (chunking normally prevents this) |
| `EBUSY: resource busy or locked` | The output file is open in Excel | Close Excel and re-run. (It otherwise auto-saves a `_NEW.xlsx` — close Excel, delete the old, rename `_NEW`) |
| `ETIMEDOUT` / `ECONNREFUSED` | Portal down or no internet | Check the portal in a browser; wait and retry |
| Zone shows `0 switches OK` | API returned empty | Retry; check the cityId is still valid |
| `Cannot parse CCMS ID` | Unexpected switch-ID format | Inspect that switch on the portal |
| `Node.js is not installed` (from the .bat) | Node missing | Install from https://nodejs.org, re-run |
| `Cannot find module 'axios'`/`exceljs` | Dependencies missing | `npm install` in the project folder (the .bat does this automatically) |

---

## Monthly routine (the short version)

1. Double-click `run-report.bat` → answer month / year / username / password → wait.
2. Double-click `run-operational-report.bat` → same answers → wait.
3. Both files appear in `Reports\<Month><Year>\`. Verify with the checklist above. Done.
