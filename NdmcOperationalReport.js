const axios = require("axios");
const ExcelJS = require("exceljs");
const https = require("https");
const readline = require("readline");
const fs = require("fs");
const path = require("path");

// ============================================================================
// NDMC OPERATIONAL HOUR REPORT
// ----------------------------------------------------------------------------
// Mirrors the manual "Operational Hour Report <Month> <Year>.xlsx" file:
// one workbook, one sheet per zone, and within each sheet ONE ROW PER DEVICE
// PER DAY from the operational endpoint (view:"2"). Columns:
//
//   A Switch Point Name | B Date | C Location | D On Hours | E OFF Hours |
//   F Output OFF Hours  | G Expected ON Hour  | H Uptime
//
// Layout notes (all verified against the reference April-2026 workbook):
//   • Durations are "HH:MM:SS" TEXT; Date is "YYYY-MM-DD" TEXT.
//   • No cell is ever left blank — a zero duration writes "00:00:00".
//   • Uptime is a FRACTION (1 = 100%) shown with the "0%" number format.
//   • Plain sheet: bold header row only, no borders/fills/freeze panes.
//
// The endpoint returns ~2 rows per device-day (one per half of the night) with
// the DAILY totals stamped on every copy, so we keep the first row per
// (device, date) rather than summing — see fetchOperationalDetailed().
//
// Reuses the same auto-login / concurrency / date-chunking approach as
// NdmcUptimeReport.js. The only data call is the operational detailed view.
// ============================================================================

// Report period. Leave 0 to be PROMPTED. Env NDMC_MONTH / NDMC_YEAR override.
let REPORT_YEAR  = Number(process.env.NDMC_YEAR)  || 0;   // 0 → ask (default: current year)
let REPORT_MONTH = Number(process.env.NDMC_MONTH) || 0;   // 0 → ask (1=Jan … 12=Dec)

// Advanced: explicit range overrides the month entirely. "YYYY-MM-DD".
const CUSTOM_START_DATE = "";
const CUSTOM_END_DATE   = "";

// Auth — same as the uptime script. Auto-login via /smartlight/login; creds from
// env NDMC_USER / NDMC_PASS or prompted. Optional: paste a JSESSIONID to skip login.
const MANUAL_JSESSIONID = "";
const PORTAL_USERNAME = process.env.NDMC_USER || "";
const PORTAL_PASSWORD = process.env.NDMC_PASS || "";

let SESSION_COOKIE = MANUAL_JSESSIONID;

// Portal 502s under load; 3 is the safe in-flight cap across all chunks.
const MAX_CONCURRENCY = 3;

const DEBUG = true;

// ============================================================================

const BASE = "https://smartlight.citilight.co:446";
const ENDPOINTS = {
    liveData:    `${BASE}/smartlight/getListViewData_v1`,
    operational: `${BASE}/VELOCITi_API/api/ccmsOperationalreportallData`,
};
const REFERERS = {
    liveData:    `${BASE}/smartlight/livedatafeed`,
    operational: `${BASE}/smartlight/operationalReport`,
};

// Tab order + names match the manual "Operational Hour Report" workbook exactly.
const ZONES = [
    { cityId: "2", cityName: "SP",          sheetName: "SP" },
    { cityId: "3", cityName: "CITY",        sheetName: "City" },
    { cityId: "4", cityName: "CIVIL LINES", sheetName: "Civil Lines" },
    { cityId: "5", cityName: "KAROL BAGH",  sheetName: "Karol Bagh" },
    { cityId: "6", cityName: "NARELA",      sheetName: "Narela" },
    { cityId: "7", cityName: "ROHINI",      sheetName: "Rohini" },
];

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const HEADERS = [
    "Switch Point Name",
    "Date",
    "Location",
    "On Hours",
    "OFF Hours",
    "Output OFF Hours",
    "Expected ON Hour",
    "Uptime",
];

const httpsAgent = new https.Agent({ keepAlive: false });

const makeHeaders = (referer) => ({
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/json",
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "*/*",
    "Origin": BASE,
    "Referer": referer,
    "Cookie": `JSESSIONID=${SESSION_COOKIE}`,
});

// ----------------------------------------------------------------------------
// Auto-login (identical approach to NdmcUptimeReport.js)
// ----------------------------------------------------------------------------
function extractJsessionid(setCookie) {
    if (!setCookie) return null;
    const list = Array.isArray(setCookie) ? setCookie : [setCookie];
    for (const c of list) {
        const m = /JSESSIONID=([^;]+)/i.exec(c);
        if (m) return m[1];
    }
    return null;
}

async function login(username, password) {
    const loginUrl = `${BASE}/smartlight/login`;
    const common = { httpsAgent, timeout: 60000, maxRedirects: 0, validateStatus: () => true };

    const seed = await axios.get(loginUrl, common);
    let sid = extractJsessionid(seed.headers["set-cookie"]);

    const form = `username=${encodeURIComponent(username)}&password=${encodeURIComponent(password)}`;
    const res = await axios.post(loginUrl, form, {
        ...common,
        headers: {
            "User-Agent": "Mozilla/5.0",
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": BASE,
            "Referer": loginUrl,
            ...(sid ? { Cookie: `JSESSIONID=${sid}` } : {}),
        },
    });
    sid = extractJsessionid(res.headers["set-cookie"]) || sid;
    if (!sid) throw new Error("login: server did not return a JSESSIONID cookie");
    return sid;
}

// True when nobody can type an answer — a scheduled task, cron, or a piped run.
// A blocked readline under Task Scheduler hangs forever holding a portal session.
const UNATTENDED = !process.stdin.isTTY || process.env.NDMC_UNATTENDED === "1";

function ask(question, { hidden = false } = {}) {
    if (UNATTENDED) {
        throw new Error(
            `Cannot prompt for "${question.trim()}" — this is an unattended run.\n` +
            `  Set the value via environment variable instead:\n` +
            `    NDMC_MONTH (1-12, or "last" for the previous month), NDMC_YEAR, NDMC_USER, NDMC_PASS`
        );
    }
    return new Promise((resolve) => {
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true });
        if (hidden) {
            rl._writeToOutput = (s) => rl.output.write(s.includes(question) ? question : "*");
        }
        rl.question(question, (answer) => {
            rl.close();
            if (hidden) process.stdout.write("\n");
            resolve(answer.trim());
        });
    });
}

async function ensureSession() {
    if (MANUAL_JSESSIONID && MANUAL_JSESSIONID !== "PASTE_YOUR_COOKIE_HERE") {
        SESSION_COOKIE = MANUAL_JSESSIONID;
        console.log("Using JSESSIONID from MANUAL_JSESSIONID (skipping login).");
        return;
    }
    const username = PORTAL_USERNAME || await ask("Portal username: ");
    const password = PORTAL_PASSWORD || await ask("Portal password: ", { hidden: true });
    if (!username || !password) throw new Error("Username and password are required to log in.");

    console.log("Logging in…");
    SESSION_COOKIE = await login(username, password);

    const test = await fetchLiveData("2");
    if (!Array.isArray(test)) {
        throw new Error("Login failed or session invalid — check the username/password.");
    }
    console.log("Login OK — session acquired.");
}

// The month a scheduled run should report on: the one that just ended. Run on
// 1–2 August and you get July; on 1–2 January you get December of the year before.
function previousMonth() {
    const now = new Date();
    const m = now.getMonth();                       // 0-11 → already the previous month in 1-12 terms
    return m === 0
        ? { month: 12, year: now.getFullYear() - 1 }
        : { month: m,  year: now.getFullYear() };
}

async function ensureReportPeriod() {
    if (CUSTOM_START_DATE && CUSTOM_END_DATE) {
        console.log(`Using custom date range ${CUSTOM_START_DATE} → ${CUSTOM_END_DATE}`);
        return;
    }

    // Unattended (scheduled) runs, or an explicit NDMC_MONTH=last, report on the
    // month that just finished — so the monthly job needs no edit each month.
    const wantsPrevious = String(process.env.NDMC_MONTH || "").toLowerCase() === "last";
    if (!(REPORT_MONTH >= 1 && REPORT_MONTH <= 12) && (wantsPrevious || UNATTENDED)) {
        const p = previousMonth();
        REPORT_MONTH = p.month;
        REPORT_YEAR  = Number(process.env.NDMC_YEAR) || p.year;
        console.log(`Unattended run — reporting on the previous month: ${MONTH_NAMES[REPORT_MONTH - 1]} ${REPORT_YEAR}`);
        return;
    }

    while (!(REPORT_MONTH >= 1 && REPORT_MONTH <= 12)) {
        const m = await ask("Which month do you want the report for? Enter 1-12 (1=Jan … 12=Dec): ");
        REPORT_MONTH = Number(m);
        if (!(REPORT_MONTH >= 1 && REPORT_MONTH <= 12)) console.log("  Please enter a whole number from 1 to 12.");
    }
    if (!(REPORT_YEAR >= 2000)) {
        const def = new Date().getFullYear();
        const y = await ask(`Which year? Press Enter for ${def}: `);
        REPORT_YEAR = y ? Number(y) : def;
        if (!(REPORT_YEAR >= 2000)) REPORT_YEAR = def;
    }
}

const monthLabel = () => `${MONTH_NAMES[REPORT_MONTH - 1]}-${String(REPORT_YEAR).slice(-2)}`;

function reportDateRange() {
    if (CUSTOM_START_DATE && CUSTOM_END_DATE) {
        return { startDate: CUSTOM_START_DATE, endDate: CUSTOM_END_DATE };
    }
    const mm = String(REPORT_MONTH).padStart(2, "0");
    const lastDay = new Date(REPORT_YEAR, REPORT_MONTH, 0).getDate();
    return {
        startDate: `${REPORT_YEAR}-${mm}-01`,
        endDate:   `${REPORT_YEAR}-${mm}-${String(lastDay).padStart(2, "0")}`,
    };
}

// Matches the manual file the portal produces, e.g. "Operational Hour Report April 2026.xlsx".
const FULL_MONTH_NAMES = ["January","February","March","April","May","June",
                          "July","August","September","October","November","December"];
const reportFilename = () => `Operational Hour Report ${FULL_MONTH_NAMES[REPORT_MONTH - 1]} ${REPORT_YEAR}.xlsx`;

// Every run's output goes into Reports/<Month><Year>/ — one folder per report month,
// created automatically if it doesn't exist. Both the Uptime and the Operational report
// for the same month land in the same folder. Returns the full path to write to.
const REPORTS_ROOT = path.join(__dirname, "Reports");
const monthFolderName = () => `${MONTH_NAMES[REPORT_MONTH - 1]}${REPORT_YEAR}`;
function reportOutputPath() {
    const dir = path.join(REPORTS_ROOT, monthFolderName());
    fs.mkdirSync(dir, { recursive: true });   // recursive: no error if it already exists
    return path.join(dir, reportFilename());
}

// First candidate key that has a usable (non-null) value wins.
function pickField(obj, candidates) {
    for (const k of candidates) if (obj && obj[k] !== undefined && obj[k] !== null && obj[k] !== "") return obj[k];
    return undefined;
}

// Seconds → "HH:MM:SS" TEXT. Also accepts a numeric string. Anything already
// containing ":" (already formatted by the API) is passed through untouched.
// A missing/zero value renders "00:00:00", never blank — the reference workbook
// has no empty cells anywhere in the grid.
function formatDuration(v) {
    if (v === undefined || v === null || v === "") return "00:00:00";
    if (typeof v === "string" && v.includes(":")) return v;
    const n = Number(v);
    if (!Number.isFinite(n)) return String(v);
    const total = Math.round(n);
    const h = Math.floor(total / 3600);
    const m = Math.floor((total % 3600) / 60);
    const s = total % 60;
    const pad = (x) => String(x).padStart(2, "0");
    return `${pad(h)}:${pad(m)}:${pad(s)}`;
}

// "2026-02-01T..." / "2026/02/01 00:00:00" → "2026-02-01".
function formatDate(v) {
    if (v === undefined || v === null || v === "") return "";
    const s = String(v).trim();
    const m = s.match(/(\d{4})[-/](\d{2})[-/](\d{2})/);
    return m ? `${m[1]}-${m[2]}-${m[3]}` : s;
}

// ----------------------------------------------------------------------------
// Concurrency + retry (same as the uptime script)
// ----------------------------------------------------------------------------
function makeSemaphore(max) {
    let active = 0;
    const queue = [];
    const acquire = () => new Promise(resolve => {
        if (active < max) { active++; resolve(); }
        else queue.push(resolve);
    });
    const release = () => {
        if (queue.length > 0) queue.shift()();
        else active--;
    };
    return { acquire, release };
}

const portalLimiter = makeSemaphore(MAX_CONCURRENCY);

async function postWithRetry(url, body, referer, label) {
    const headers = makeHeaders(referer);
    const opts = { headers, httpsAgent, timeout: 300000 };
    await portalLimiter.acquire();
    try {
        try {
            return (await axios.post(url, body, opts)).data;
        } catch (err) {
            console.log(`    [retry] ${label}: ${err.message}`);
            await new Promise(r => setTimeout(r, 3000));
            return (await axios.post(url, body, opts)).data;
        }
    } finally {
        portalLimiter.release();
    }
}

const fetchLiveData = (cityId) => postWithRetry(
    ENDPOINTS.liveData,
    { cityId, deviceType: "1", zoneName: "0", wardName: "0", streetName: "0", userId: "10" },
    REFERERS.liveData,
    "liveData",
);

// Detailed operational view (view:"2"), CCMS (deviceType:"1"), all devices in the city.
const fetchOperationalSingle = (cityId, startDate, endDate) => postWithRetry(
    ENDPOINTS.operational,
    { cityId, deviceType: "1", startDate, endDate, deviceArray: [], zoneName: "all", wardName: "all", view: "2" },
    REFERERS.operational,
    "operational",
);

function chunkDateRange(startDate, endDate, chunkDays) {
    const out = [];
    const end = new Date(endDate + "T00:00:00Z");
    let cursor = new Date(startDate + "T00:00:00Z");
    while (cursor <= end) {
        const chunkEnd = new Date(cursor);
        chunkEnd.setUTCDate(chunkEnd.getUTCDate() + chunkDays - 1);
        if (chunkEnd > end) chunkEnd.setTime(end.getTime());
        out.push({ sd: cursor.toISOString().slice(0, 10), ed: chunkEnd.toISOString().slice(0, 10) });
        cursor = new Date(chunkEnd);
        cursor.setUTCDate(cursor.getUTCDate() + 1);
    }
    return out;
}

// ----------------------------------------------------------------------------
// Field mapping. The detailed-view field names aren't fully documented — only
// device_name/updated_on/expected_on/output_off are confirmed (used by the uptime
// script). For everything else we try several candidate keys, and DEBUG prints the
// real keys of the first row so any unmapped column can be fixed in one place.
// ----------------------------------------------------------------------------
// Confirmed from the live detailed-view response (see probe output 2026-06):
//   actual_on/off_seconds, expected_on, output_off = duration seconds, and they are
//   DAILY totals repeated on every copy of a device-day. starttime/endtime/powercut_*
//   are epoch seconds — no longer used, the report has no timestamp columns.
const FIELD = {
    switchPoint:   ["device_name"],
    location:      ["switch_location", "location"],
    date:          ["updated_on"],
    totalOn:       ["actual_on_seconds"],
    outputOff:     ["output_off"],
    totalOff:      ["actual_off_seconds"],
    expectedOn:    ["expected_on"],
    uptime:        ["uptime"],
};

// Build one Excel row [A..H] from a raw operational detail record. liveLoc is the
// device→location map from the live feed, used when the row carries no location.
function toRow(r, liveLoc) {
    const sp   = pickField(r, FIELD.switchPoint);
    const loc  = pickField(r, FIELD.location) ?? (sp ? liveLoc.get(String(sp)) : undefined);
    return [
        sp ?? "",
        formatDate(pickField(r, FIELD.date)),
        loc ?? "",
        formatDuration(pickField(r, FIELD.totalOn)),
        formatDuration(pickField(r, FIELD.totalOff)),
        formatDuration(pickField(r, FIELD.outputOff)),
        formatDuration(pickField(r, FIELD.expectedOn)),
        toUptimeFraction(pickField(r, FIELD.uptime)),
    ];
}

// Uptime as a FRACTION for the "0%" number format: the API reports a percentage
// (100 = fully up), the sheet stores 1. Missing → 0, so no cell is left blank.
function toUptimeFraction(v) {
    if (v === undefined || v === null || v === "") return 0;
    const n = Number(v);
    return Number.isFinite(n) ? n / 100 : String(v);
}

// Fetch the whole month of detailed rows for one city, de-duplicating exact repeats
// as they stream in (the API repeats records, so we key each unique segment once).
async function fetchOperationalDetailed(cityId, startDate, endDate, chunkDays = 4) {
    const chunks = chunkDateRange(startDate, endDate, chunkDays);
    const byKey = new Map();   // composite key → raw row (dedups the ~Nx repeats)
    let firstRowKeys = null, totalRows = 0, failedChunks = 0;
    await Promise.all(chunks.map(async ({ sd, ed }) => {
        console.log(`    op chunk ${sd} → ${ed}`);
        let part;
        try { part = await fetchOperationalSingle(cityId, sd, ed); }
        catch (err) { console.log(`      op chunk ${sd} failed: ${err.message}`); failedChunks++; return; }
        if (!Array.isArray(part)) { failedChunks++; return; }
        if (!firstRowKeys && part[0]) firstRowKeys = Object.keys(part[0]);
        for (let i = 0; i < part.length; i++) {
            const r = part[i];
            const sp   = pickField(r, FIELD.switchPoint);
            const date = pickField(r, FIELD.date);
            if (!sp) continue;
            // ONE row per (device, date). The endpoint repeats each device-day many
            // times — once per half of the night, each copy repeated further — but
            // actual_on/off_seconds, output_off and expected_on are DAILY totals
            // stamped identically on every copy. So we keep the FIRST copy rather than
            // summing; summing would double every duration. On a sunset/sunrise
            // schedule-change day the API returns differing variants, and keeping the
            // first preserves the pre-change (older-schedule) values, matching the
            // manual report.
            const key = `${sp}|${date}`;
            if (!byKey.has(key)) byKey.set(key, r);
        }
        totalRows += part.length;
    }));
    return { rows: [...byKey.values()], firstRowKeys, totalRows, uniqueRows: byKey.size, failedChunks };
}

async function buildZoneRows(zone, dateRange) {
    console.log(`\n[${zone.sheetName}] cityId=${zone.cityId}`);

    // Live feed → device→location map (fallback for column B) and a session check.
    const liveData = await fetchLiveData(zone.cityId);
    const liveLoc = new Map();
    for (const r of (liveData || [])) {
        const id = pickField(r, FIELD.switchPoint);
        const loc = pickField(r, FIELD.location);
        if (id && loc) liveLoc.set(String(id), loc);
    }
    console.log(`  live data: ${(liveData || []).length} switches (${liveLoc.size} with location)`);

    const agg = await fetchOperationalDetailed(zone.cityId, dateRange.startDate, dateRange.endDate);
    if (DEBUG && agg.firstRowKeys) {
        console.log("  >>> operational detail row KEYS:", agg.firstRowKeys);
        if (agg.rows[0]) console.log("  >>> operational detail row SAMPLE:", JSON.stringify(agg.rows[0]).slice(0, 600));
    }
    console.log(`  operational: ${agg.totalRows} raw rows → ${agg.uniqueRows} device-days (${agg.failedChunks} failed chunks)`);

    // Build, then sort by Switch Point, then Date — matches the manual file.
    const rows = agg.rows.map(r => toRow(r, liveLoc));
    rows.sort((a, b) =>
        String(a[0]).localeCompare(String(b[0])) ||
        String(a[1]).localeCompare(String(b[1]))
    );
    return rows;
}

// ----------------------------------------------------------------------------
// Excel layout (8 columns, matching the manual "Operational Hour Report" file).
// Widths are copied verbatim from the reference workbook's SP sheet. The
// reference is a plain grid — no borders, no fills, no freeze pane, default
// alignment — so we only set widths, the bold header, and Uptime's "0%" format.
// ----------------------------------------------------------------------------
const COLUMN_SPEC = [
    { width: 22.14 },                  // A  Switch Point Name
    { width: 13    },                  // B  Date
    { width: 52    },                  // C  Location
    { width: 10.43 },                  // D  On Hours
    { width: 11.71 },                  // E  OFF Hours
    { width: 20.86 },                  // F  Output OFF Hours
    { width: 20.86 },                  // G  Expected ON Hour
    { width: 9.14, numFmt: "0%" },     // H  Uptime (fraction: 1 = 100%)
];

const NCOLS = COLUMN_SPEC.length;

function buildSheet(workbook, zone, rows) {
    const ws = workbook.addWorksheet(zone.sheetName);

    COLUMN_SPEC.forEach((spec, i) => { ws.getColumn(i + 1).width = spec.width; });

    // Row 1: bold header, otherwise unstyled.
    ws.addRow(HEADERS);
    ws.getRow(1).eachCell({ includeEmpty: true }, (cell, c) => {
        if (c > NCOLS) return;
        cell.font = { bold: true };
    });

    for (const r of rows) {
        const row = ws.addRow(r);
        row.eachCell({ includeEmpty: true }, (cell, c) => {
            if (c > NCOLS) return;
            const spec = COLUMN_SPEC[c - 1];
            if (spec.numFmt && typeof cell.value === "number") cell.numFmt = spec.numFmt;
        });
    }
    return ws;
}

async function main() {
    console.log("=========================================");
    console.log("   NDMC Operational Hour Report");
    console.log("=========================================\n");

    await ensureReportPeriod();
    const dateRange = reportDateRange();
    console.log(`\nReport: ${monthLabel()}   (${dateRange.startDate} → ${dateRange.endDate})\n`);

    await ensureSession();

    const workbook = new ExcelJS.Workbook();
    const summary = [];

    for (const zone of ZONES) {
        try {
            const rows = await buildZoneRows(zone, dateRange);
            buildSheet(workbook, zone, rows);
            summary.push({ zone: zone.sheetName, rows: rows.length, status: "OK" });
            await new Promise(r => setTimeout(r, 2000));
        } catch (err) {
            console.error(`  [${zone.sheetName}] FAILED: ${err.message}`);
            summary.push({ zone: zone.sheetName, rows: 0, status: `FAIL: ${err.message}` });
            buildSheet(workbook, zone, []);
        }
    }

    const filename = reportOutputPath();   // Reports/<Month><Year>/Operational Hour Report <Month> <Year>.xlsx
    let written = filename;
    try {
        await workbook.xlsx.writeFile(filename);
    } catch (err) {
        if (err.code === "EBUSY" || err.code === "EPERM") {
            const fallback = filename.replace(/\.xlsx$/, `_NEW.xlsx`);
            await workbook.xlsx.writeFile(fallback);
            written = fallback;
            console.log(`\n⚠️  "${filename}" was open in Excel — saved to "${fallback}" instead.`);
            console.log(`    Close Excel, delete the old file, and rename "${fallback}" → "${filename}".`);
        } else {
            throw err;
        }
    }

    console.log("\n--- Summary ---");
    for (const s of summary) console.log(`  ${s.zone.padEnd(12)} ${String(s.rows).padStart(6)} rows  ${s.status}`);
    console.log(`\nWritten: ${written}`);

    if (process.platform === "win32") {
        try { require("child_process").exec(`start "" "${written}"`); } catch { /* ignore */ }
    }
}

main().catch((err) => {
    console.error(`\n❌ ${err.message}`);
    process.exitCode = 1;
});
