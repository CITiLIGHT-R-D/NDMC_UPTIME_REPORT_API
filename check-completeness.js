// ============================================================================
// COMPLETENESS CHECK — run this on a generated Uptime report before sending it.
// ----------------------------------------------------------------------------
// Catches the failure mode that produced the bad July 2026 Rohini sheet: the
// portal drops chunks under load, the generator logs it but still writes the
// zone as "OK", and the resulting sheet holds only part of the month.
//
// Column E (Night Duration / Estimated time of Operation) is the tell. Every
// zone sees the same sunset-to-sunrise schedule, so E must be effectively
// IDENTICAL across all six zones. A zone that lost chunks has a proportionally
// smaller E — July 2026 Rohini showed 40.67 against 315.17 elsewhere, i.e.
// 4 days out of 31.
//
// Usage:
//   node check-completeness.js                                  (newest report)
//   node check-completeness.js "Reports/Jul2026/NDMC_UptimeReport_Jul2026.xlsx"
//
// Exit code 0 = all zones complete, 1 = at least one zone is short.
// ============================================================================

const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// A zone is flagged when its E differs from the group's by more than this.
// The zones genuinely share one schedule, so real spread is ~0.
const TOLERANCE = 0.01;   // 1%

// Find the most recently modified NDMC_UptimeReport_*.xlsx under Reports/.
function newestReport() {
    const root = path.join(__dirname, "Reports");
    if (!fs.existsSync(root)) return null;
    const hits = [];
    for (const dir of fs.readdirSync(root)) {
        const full = path.join(root, dir);
        if (!fs.statSync(full).isDirectory()) continue;
        for (const f of fs.readdirSync(full)) {
            if (/^NDMC_UptimeReport_.*\.xlsx$/i.test(f) && !f.startsWith("~$")) {
                const p = path.join(full, f);
                hits.push({ p, m: fs.statSync(p).mtimeMs });
            }
        }
    }
    hits.sort((a, b) => b.m - a.m);
    return hits.length ? hits[0].p : null;
}

// "NDMC_UptimeReport_Jul2026.xlsx" → days in that month, or null.
function daysInMonthFromName(file) {
    const m = path.basename(file).match(/_([A-Za-z]{3})(\d{4})\.xlsx$/);
    if (!m) return null;
    const idx = MONTH_NAMES.findIndex(n => n.toLowerCase() === m[1].toLowerCase());
    if (idx < 0) return null;
    return { days: new Date(Number(m[2]), idx + 1, 0).getDate(), label: `${m[1]} ${m[2]}` };
}

const median = (xs) => {
    const s = [...xs].sort((a, b) => a - b);
    const mid = s.length >> 1;
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
};

(async () => {
    const file = process.argv[2] ? path.resolve(process.argv[2]) : newestReport();
    if (!file || !fs.existsSync(file)) {
        console.error("No report found. Pass a path, e.g.:\n  node check-completeness.js \"Reports/Jul2026/NDMC_UptimeReport_Jul2026.xlsx\"");
        process.exit(1);
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);

    const period = daysInMonthFromName(file);
    console.log("=".repeat(78));
    console.log("  COMPLETENESS CHECK");
    console.log("=".repeat(78));
    console.log(`  File:   ${file}`);
    if (period) console.log(`  Period: ${period.label} (${period.days} days)`);
    console.log("");

    // Collect column E per zone. Data starts at row 3 (row 1 title, row 2 header).
    const zones = [];
    for (const ws of wb.worksheets) {
        const eVals = new Set();
        let rows = 0;
        for (let r = 3; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            if (!row.getCell(2).value) continue;       // no Switch ID -> not a data row
            rows++;
            const e = Number(row.getCell(5).value);
            if (Number.isFinite(e)) eVals.add(e);
        }
        zones.push({ name: ws.name, rows, eVals: [...eVals] });
    }

    const withData = zones.filter(z => z.rows > 0 && z.eVals.length);
    if (!withData.length) {
        console.error("  FAIL: no data rows found in any sheet.");
        process.exit(1);
    }

    // The group's true E: the median across zones. Robust even if one or two are short.
    const groupE = median(withData.map(z => z.eVals[0]));
    const nightHours = period ? groupE / period.days : null;

    console.log("  Zone            Rows   Column E    vs group   Implied days   Status");
    console.log("  " + "-".repeat(74));

    let bad = 0;
    for (const z of zones) {
        if (!z.rows) {
            console.log(`  ${z.name.padEnd(14)} ${String(z.rows).padStart(5)}   ${"-".padStart(9)}   ${"-".padStart(8)}   ${"-".padStart(12)}   EMPTY`);
            bad++;
            continue;
        }
        if (z.eVals.length > 1) {
            console.log(`  ${z.name.padEnd(14)} ${String(z.rows).padStart(5)}   ${z.eVals.length} DISTINCT E VALUES — expected exactly 1`);
            bad++;
            continue;
        }
        const e = z.eVals[0];
        const ratio = groupE ? e / groupE : 0;
        const days = nightHours ? e / nightHours : null;
        const ok = Math.abs(1 - ratio) <= TOLERANCE;
        if (!ok) bad++;
        console.log(
            `  ${z.name.padEnd(14)} ${String(z.rows).padStart(5)}   ${e.toFixed(4).padStart(9)}   ${(ratio * 100).toFixed(1).padStart(7)}%   ${(days === null ? "-" : days.toFixed(1)).padStart(12)}   ${ok ? "OK" : "*** INCOMPLETE ***"}`
        );
    }

    console.log("");
    if (bad) {
        console.log(`  RESULT: FAIL — ${bad} zone(s) look incomplete.`);
        console.log("");
        console.log("  A short column E means the generator lost date-chunks for that zone.");
        console.log("  The data is on the portal; the download was partial. Re-run the report");
        console.log("  for that month when the portal is under light load, then re-check.");
        console.log("");
        console.log("  Do NOT send the report until every zone reads OK — the hours and kWh");
        console.log("  columns (E, F, J, K) are understated for any zone flagged above.");
        process.exit(1);
    }

    console.log(`  RESULT: PASS — all ${zones.length} zones agree on column E (${groupE.toFixed(4)}).`);
    if (nightHours) console.log(`  Implies ${period.days} nights at ~${nightHours.toFixed(4)} h/night.`);
    process.exit(0);
})().catch((err) => {
    console.error(`\nERROR: ${err.message}`);
    process.exit(1);
});
