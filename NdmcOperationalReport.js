const axios = require("axios");
const ExcelJS = require("exceljs");
const https = require("https");
const readline = require("readline");
const fs = require("fs");
const path = require("path");

// ============================================================================
// NDMC OPERATIONAL (DETAILED) REPORT
// ----------------------------------------------------------------------------
// Mirrors the manual "Operation_uptime_Reports_<Month><Year>.xlsx" file:
// one workbook, one sheet per zone, and within each sheet the RAW detailed-view
// rows from the operational endpoint (view:"2") — one row per device per
// night-segment. No aggregation, no computed business columns. Columns:
//
//   A Switch Point | B Location | C Date | D Start Time | E End Time |
//   F Total On Hours | G Start Time | H End Time | I Output OFF |
//   J Total Off Hours | K Expected ON Hour | L Uptime %
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

// Tab order + names match the manual Operation_uptime_Reports workbook exactly.
const ZONES = [
    { cityId: "2", cityName: "SP",          sheetName: "SP" },
    { cityId: "3", cityName: "CITY",        sheetName: "City" },
    { cityId: "5", cityName: "KAROL BAGH",  sheetName: "Karol Bhag" },
    { cityId: "4", cityName: "CIVIL LINES", sheetName: "Civil Line" },
    { cityId: "6", cityName: "NARELA",      sheetName: "Narela" },
    { cityId: "7", cityName: "ROHINI",      sheetName: "Rohini" },
];

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const HEADERS = [
    "Switch Point",
    "Location",
    "Date",
    "Start Time",
    "End Time",
    "Total On Hours",
    "Start Time",
    "End Time",
    "Output OFF Duration",
    "Total Off Hours",
    "Expected ON Hour",
    "Uptime %",
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

function ask(question, { hidden = false } = {}) {
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

async function ensureReportPeriod() {
    if (CUSTOM_START_DATE && CUSTOM_END_DATE) {
        console.log(`Using custom date range ${CUSTOM_START_DATE} → ${CUSTOM_END_DATE}`);
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

const reportFilename = () => `Operation_uptime_Reports_${MONTH_NAMES[REPORT_MONTH - 1]}${REPORT_YEAR}.xlsx`;

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

// Seconds → "HH:MM:SS". Also accepts a numeric string. Anything already containing
// ":" (already formatted by the API) is passed through untouched. With
// { blankZero:true } a value of 0 renders blank (used for Output OFF).
function formatDuration(v, { blankZero = false } = {}) {
    if (v === undefined || v === null || v === "") return "";
    if (typeof v === "string" && v.includes(":")) return v;
    const n = Number(v);
    if (!Number.isFinite(n)) return String(v);
    if (blankZero && n === 0) return "";
    const total = Math.round(n);
    const h = Math.floor(total / 3600);
    const m = Math.floor((total % 3600) / 60);
    const s = total % 60;
    const pad = (x) => String(x).padStart(2, "0");
    return `${pad(h)}:${pad(m)}:${pad(s)}`;
}

// Epoch seconds → IST (UTC+5:30) "YYYY/MM/DD HH:MM:SS". Null/blank → "".
// (starttime/endtime/powercut_* arrive as Unix-epoch seconds.) Shifting by 19800s
// then reading the UTC parts makes the result independent of the machine timezone.
function formatEpoch(v) {
    if (v === undefined || v === null || v === "") return "";
    const n = Number(v);
    if (!Number.isFinite(n)) return String(v);
    const d = new Date((n + 19800) * 1000);
    const p = (x) => String(x).padStart(2, "0");
    return `${d.getUTCFullYear()}/${p(d.getUTCMonth() + 1)}/${p(d.getUTCDate())} `
         + `${p(d.getUTCHours())}:${p(d.getUTCMinutes())}:${p(d.getUTCSeconds())}`;
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
//   starttime/endtime/powercut_* = epoch seconds; actual_on/off_seconds, expected_on,
//   output_off = duration seconds; ontime/offtime are epochs (NOT durations) — unused.
const FIELD = {
    switchPoint:   ["device_name"],
    location:      ["switch_location", "location"],
    date:          ["updated_on"],
    onStart:       ["starttime"],
    onEnd:         ["endtime"],
    totalOn:       ["actual_on_seconds"],
    offStart:      ["powercut_start"],
    offEnd:        ["powercut_end"],
    outputOff:     ["output_off"],
    totalOff:      ["actual_off_seconds"],
    expectedOn:    ["expected_on"],
    uptime:        ["uptime"],
};

// Build one Excel row [A..L] from a raw operational detail record. liveLoc is the
// device→location map from the live feed, used when the row carries no location.
function toRow(r, liveLoc) {
    const sp   = pickField(r, FIELD.switchPoint);
    const loc  = pickField(r, FIELD.location) ?? (sp ? liveLoc.get(String(sp)) : undefined);
    return [
        sp ?? "",
        loc ?? "",
        formatDate(pickField(r, FIELD.date)),
        formatEpoch(pickField(r, FIELD.onStart)),
        formatEpoch(pickField(r, FIELD.onEnd)),
        formatDuration(pickField(r, FIELD.totalOn)),
        formatEpoch(pickField(r, FIELD.offStart)),
        formatEpoch(pickField(r, FIELD.offEnd)),
        formatDuration(pickField(r, FIELD.outputOff), { blankZero: true }),
        formatDuration(pickField(r, FIELD.totalOff)),
        formatDuration(pickField(r, FIELD.expectedOn)),
        toUptimeNumber(pickField(r, FIELD.uptime)),
    ];
}

// Uptime as a 2-decimal number (the sheet shows 100.00). Null/blank → "".
function toUptimeNumber(v) {
    if (v === undefined || v === null || v === "") return "";
    const n = Number(v);
    return Number.isFinite(n) ? Number(n.toFixed(2)) : String(v);
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
            const sp    = pickField(r, FIELD.switchPoint);
            const date  = pickField(r, FIELD.date);
            const start = pickField(r, FIELD.onStart);
            if (!sp) continue;
            // Collapse to one segment per (device, date, half-of-night): AM (after-
            // midnight, starts before noon IST) and PM (evening). The API repeats each
            // record many times; on a sunset/sunrise schedule-change day it returns TWO
            // differing variants per half, so a start/end-based key keeps both and inflates
            // the row count (~+2 rows per device on that day). Keeping the FIRST variant per
            // half matches the manual report: 2 rows per device per day, and the pre-change
            // (older-schedule) variant is kept on the transition day.
            const startSec = Number(start);
            const half = Number.isFinite(startSec) && ((startSec + 19800) % 86400) / 3600 >= 12 ? "PM" : "AM";
            const key = `${sp}|${date}|${half}`;
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
    console.log(`  operational: ${agg.totalRows} raw rows → ${agg.uniqueRows} unique segments (${agg.failedChunks} failed chunks)`);

    // Build, then sort by Switch Point, then Date, then Start Time — matches the manual file.
    const rows = agg.rows.map(r => toRow(r, liveLoc));
    rows.sort((a, b) =>
        String(a[0]).localeCompare(String(b[0])) ||
        String(a[2]).localeCompare(String(b[2])) ||
        String(a[3]).localeCompare(String(b[3]))
    );
    return rows;
}

// ----------------------------------------------------------------------------
// Excel layout (12 columns, matching the manual Operation_uptime_Reports file)
// ----------------------------------------------------------------------------
const COLUMN_SPEC = [
    { width: 16, align: "left"   },   // A  Switch Point
    { width: 46, align: "left"   },   // B  Location
    { width: 12, align: "center" },   // C  Date
    { width: 20, align: "center" },   // D  Start Time
    { width: 20, align: "center" },   // E  End Time
    { width: 14, align: "center" },   // F  Total On Hours
    { width: 14, align: "center" },   // G  Start Time
    { width: 14, align: "center" },   // H  End Time
    { width: 12, align: "center" },   // I  Output OFF
    { width: 14, align: "center" },   // J  Total Off Hours
    { width: 14, align: "center" },   // K  Expected ON Hour
    { width: 10, align: "right", numFmt: "0.00" },   // L  Uptime %
];

const THIN = { style: "thin", color: { argb: "FF000000" } };
const ALL_BORDERS = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const NCOLS = COLUMN_SPEC.length;

function buildSheet(workbook, zone, rows) {
    const ws = workbook.addWorksheet(zone.sheetName, {
        views: [{ state: "frozen", xSplit: 0, ySplit: 1, topLeftCell: "A2", activeCell: "A2" }],
    });

    COLUMN_SPEC.forEach((spec, i) => { ws.getColumn(i + 1).width = spec.width; });

    // Row 1: bold header.
    ws.addRow(HEADERS);
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell({ includeEmpty: true }, (cell, c) => {
        if (c > NCOLS) return;
        cell.font = { bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.border = ALL_BORDERS;
    });

    for (const r of rows) {
        const row = ws.addRow(r);
        row.eachCell({ includeEmpty: true }, (cell, c) => {
            if (c > NCOLS) return;
            const spec = COLUMN_SPEC[c - 1];
            cell.alignment = { horizontal: spec.align, vertical: "middle" };
            if (spec.numFmt && typeof cell.value === "number") cell.numFmt = spec.numFmt;
            cell.border = ALL_BORDERS;
        });
    }
    return ws;
}

async function main() {
    console.log("=========================================");
    console.log("   NDMC Operational (Detailed) Report");
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

    const filename = reportOutputPath();   // Reports/<Month><Year>/Operation_uptime_Reports_....xlsx
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
