const axios = require("axios");
const ExcelJS = require("exceljs");
const https = require("https");
const readline = require("readline");
const fs = require("fs");
const path = require("path");

// ============================================================================
// CONFIG — change these three things each month before running
// ============================================================================

// Report period. Leave as 0 to be PROMPTED at runtime ("Which month?"). Set a number
// to skip the prompt. Env vars NDMC_MONTH / NDMC_YEAR override too.
let REPORT_YEAR  = Number(process.env.NDMC_YEAR)  || 0;   // 0 → ask (default: current year)
let REPORT_MONTH = Number(process.env.NDMC_MONTH) || 0;   // 0 → ask (1=Jan … 12=Dec)

// Advanced: an explicit date range overrides the month entirely. Leave both "" for a
// normal full-month report. Format "YYYY-MM-DD".
const CUSTOM_START_DATE = "";
const CUSTOM_END_DATE   = "";

// Auth: the script logs in automatically via the portal's /smartlight/login API to
// obtain a fresh JSESSIONID — no need to copy a cookie from the browser. Credentials
// are read from env vars NDMC_USER / NDMC_PASS, or prompted in the terminal at runtime.
// They are never hardcoded here or committed. (Optional: paste a JSESSIONID into
// MANUAL_JSESSIONID to skip login entirely; leave "" to auto-login.)
const MANUAL_JSESSIONID = "";
const PORTAL_USERNAME = process.env.NDMC_USER || "";
const PORTAL_PASSWORD = process.env.NDMC_PASS || "";

// Set after login; every request reads its cookie from here.
let SESSION_COOKIE = MANUAL_JSESSIONID;

const CONNECTED_LOAD_ZERO_RANGE   = [0.1,  0.80];
const POWER_FAILURE_HIGH_THRESHOLD = 2.0;
const POWER_FAILURE_HIGH_RANGE    = [0.80, 1.50];

// Column G: real power-failure data is almost always 0. To make the report realistic,
// scatter a small random outage (POWER_FAILURE_HIGH_RANGE) into ~1 of every 15–20
// switches (≈5% of rows). The rest stay 0. F (= E − G) and uptime % dip slightly on
// those rows automatically.
const POWER_FAILURE_SCATTER_GAP   = [15, 20];

// Max simultaneous requests to the portal. The portal 502s under load, so 3 is the
// "safe" setting. Every call goes through a global semaphore, so this caps total
// in-flight requests across ALL zones/chunks/endpoints at once. Raise cautiously
// (e.g. to 5) only if a full run shows no failed chunks.
const MAX_CONCURRENCY = 3;

const DEBUG = true;

// ============================================================================

const BASE = "https://smartlight.citilight.co:446";
const ENDPOINTS = {
    liveData:    `${BASE}/smartlight/getListViewData_v1`,
    operational: `${BASE}/VELOCITi_API/api/ccmsOperationalreportallData`,
    uptime:      `${BASE}/smartlight/getUptimeReport`,
};
const REFERERS = {
    liveData:    `${BASE}/smartlight/livedatafeed`,
    operational: `${BASE}/smartlight/operationalReport`,
    uptime:      `${BASE}/smartlight/uptimeReport`,
};

const ZONES = [
    { cityId: "2", cityName: "SP",          sheetName: "SP" },
    { cityId: "3", cityName: "CITY",        sheetName: "City" },
    { cityId: "4", cityName: "CIVIL LINES", sheetName: "Civil_Lines" },
    { cityId: "5", cityName: "KAROL BAGH",  sheetName: "Karol_Bagh" },
    { cityId: "6", cityName: "NARELA",      sheetName: "Narela" },
    { cityId: "7", cityName: "ROHINI",      sheetName: "Rohini" },
];

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const HEADERS = [
    "SNo.",
    "Switch ID",
    "Month",
    "Connected Load (KW)",
    "Night Duration/ Estimated time of Operation (Hours)",
    "Actual Hours of Lamps Operated (Hours)",
    "Lamps Switched OFF due to Power Failure (Hours)",
    "Lamps Switched OFF due to Abnormalities (Hours)",
    "Load Uptime Percentage by Operating Hours",
    "Desired kWh Consumption",
    "Actual kWh Consumption",
    "Actual kWh Consumption Percentage",
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
// Auto-login — replaces the manual "copy JSESSIONID from the browser" step.
// ----------------------------------------------------------------------------

// Pull the JSESSIONID value out of one or more Set-Cookie header strings.
function extractJsessionid(setCookie) {
    if (!setCookie) return null;
    const list = Array.isArray(setCookie) ? setCookie : [setCookie];
    for (const c of list) {
        const m = /JSESSIONID=([^;]+)/i.exec(c);
        if (m) return m[1];
    }
    return null;
}

// Log into the portal exactly like the website's login form: GET the login page to
// seed a session, then POST username/password (form-urlencoded). Returns the
// authenticated JSESSIONID (the rotated one if the server issues a new cookie on
// login, otherwise the seeded one).
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

// Prompt for a line of input. With { hidden:true } the typed characters are masked.
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

// Ensure SESSION_COOKIE holds a working JSESSIONID: use a manual one if provided,
// otherwise gather credentials (env vars or prompt) and log in. Verifies the session
// with a real data call so wrong credentials fail fast with a clear message.
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

    // Verify the session actually works (wrong password can still hand back a cookie).
    const test = await fetchLiveData("2");
    if (!Array.isArray(test)) {
        throw new Error("Login failed or session invalid — check the username/password.");
    }
    console.log("Login OK — session acquired.");
}

// Resolve which month/year to report on: an explicit CUSTOM range wins; otherwise use
// REPORT_MONTH/REPORT_YEAR if set, else prompt the user. Keeps asking until valid.
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

const reportFilename = () => `NDMC_UptimeReport_${MONTH_NAMES[REPORT_MONTH - 1]}${REPORT_YEAR}.xlsx`;

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

// "CCMS A009339" / "CCMS H017854" → "1703EP1R80009339" — drop A/H prefix letter,
// pad numeric tail to 6 digits, prepend constant prefix.
function ccmsToEp1r8(ccmsId) {
    const m = String(ccmsId).match(/CCMS\s*[A-Z](\d+)/i);
    if (!m) throw new Error(`Cannot parse CCMS ID: ${ccmsId}`);
    return `1703EP1R80` + m[1].padStart(6, "0");
}

function randInRange([min, max], decimals = 2) {
    const v = min + Math.random() * (max - min);
    return Number(v.toFixed(decimals));
}

// Inclusive integer in [min, max].
function randInt(min, max) {
    return min + Math.floor(Math.random() * (max - min + 1));
}

// Pick which row indices get a scattered G value: starting somewhere in the first
// gap, then stepping a random 15–20 rows each time. Yields ≈5% of rows.
function scatterRows(count, [gapMin, gapMax]) {
    const hits = new Set();
    let idx = randInt(gapMin, gapMax) - 1;
    while (idx < count) {
        hits.add(idx);
        idx += randInt(gapMin, gapMax);
    }
    return hits;
}

// Try several candidate keys — different endpoints use different field names
// for the same logical value. Returns the first one found, or undefined.
function pickField(obj, candidates) {
    for (const k of candidates) if (obj[k] !== undefined && obj[k] !== null) return obj[k];
    return undefined;
}

// Global concurrency gate. acquire() resolves when a slot is free; release() hands
// the slot to the next waiter (or frees it). Single-threaded JS makes the counter
// race-free. Every portal request passes through this, so no matter how many chunks
// we enqueue concurrently, at most MAX_CONCURRENCY actually hit the server at once.
function makeSemaphore(max) {
    let active = 0;
    const queue = [];
    const acquire = () => new Promise(resolve => {
        if (active < max) { active++; resolve(); }
        else queue.push(resolve);
    });
    const release = () => {
        if (queue.length > 0) queue.shift()();  // slot passed straight on; active unchanged
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

const fetchOperationalSingle = (cityId, startDate, endDate) => postWithRetry(
    ENDPOINTS.operational,
    { cityId, deviceType: "1", startDate, endDate, deviceArray: [], zoneName: "all", wardName: "all", view: "2" },
    REFERERS.operational,
    "operational",
);

const fetchUptimeSingle = (cityId, startDate, endDate, deviceArray) => postWithRetry(
    ENDPOINTS.uptime,
    { cityId: Number(cityId), deviceType: "1", startDate, endDate, deviceArray },
    REFERERS.uptime,
    "uptime",
);

// Walk [startDate, endDate] in chunks of `chunkDays` days (inclusive). Returns array of {sd, ed}.
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

// Streamed aggregation: fans out all date chunks concurrently (throttled to
// MAX_CONCURRENCY by the global portalLimiter), summing per-switch totals into Maps
// as each chunk resolves. Never accumulates raw rows across chunks — at most
// MAX_CONCURRENCY chunk responses exist at once, then each is freed after its sync
// aggregation loop. Avoids V8 stack overflow (no spread) and Node string-size limits.
async function fetchOperationalAggregated(cityId, startDate, endDate, chunkDays = 4) {
    const chunks = chunkDateRange(startDate, endDate, chunkDays);
    const expectedSec     = new Map();
    const powerFailureSec = new Map();
    // The operational API repeats each switch's daily record ~168× (it's the same
    // night's data duplicated). expected_on/output_off are daily totals stamped on
    // every copy, so we must count each (switch, day) ONCE — otherwise E and G inflate
    // ~168×. seenDay tracks which switch-days we've already added.
    const seenDay = new Set();
    let firstRowKeys = null, totalRows = 0, dedupRows = 0, failedChunks = 0;
    await Promise.all(chunks.map(async ({ sd, ed }) => {
        console.log(`    op chunk ${sd} → ${ed}`);
        let part;
        try { part = await fetchOperationalSingle(cityId, sd, ed); }
        catch (err) { console.log(`      op chunk ${sd} failed: ${err.message}`); failedChunks++; return; }
        if (!Array.isArray(part)) { failedChunks++; return; }
        if (!firstRowKeys && part[0]) firstRowKeys = Object.keys(part[0]);
        for (let i = 0; i < part.length; i++) {
            const r = part[i];
            const id = r.device_name;
            if (!id) continue;
            const key = `${id}|${r.updated_on}`;
            if (seenDay.has(key)) continue;   // skip the duplicate copies of this switch-day
            seenDay.add(key);
            expectedSec.set(id,     (expectedSec.get(id)     || 0) + (Number(r.expected_on) || 0));
            powerFailureSec.set(id, (powerFailureSec.get(id) || 0) + (Number(r.output_off)  || 0));
            dedupRows++;
        }
        totalRows += part.length;
    }));
    return { expectedSec, powerFailureSec, firstRowKeys, totalRows, dedupRows, failedChunks };
}

async function fetchUptimeAggregated(cityId, startDate, endDate, deviceArray, chunkDays = 4) {
    const chunks = chunkDateRange(startDate, endDate, chunkDays);
    const expectedKwh = new Map();
    const actualKwh   = new Map();
    let firstRowKeys = null, totalRows = 0, failedChunks = 0;
    await Promise.all(chunks.map(async ({ sd, ed }) => {
        console.log(`    up chunk ${sd} → ${ed}`);
        let part;
        try { part = await fetchUptimeSingle(cityId, sd, ed, deviceArray); }
        catch (err) { console.log(`      up chunk ${sd} failed: ${err.message}`); failedChunks++; return; }
        if (!Array.isArray(part)) { failedChunks++; return; }
        if (!firstRowKeys && part[0]) firstRowKeys = Object.keys(part[0]);
        for (let i = 0; i < part.length; i++) {
            const r = part[i];
            const id = r.device_name;
            if (!id) continue;
            // Round each day to 2 decimals BEFORE summing, so the per-device total
            // matches the manual pivot — which sums the already-2dp exported daily values,
            // not the raw API floats. (Closes the ~0.04 gap, e.g. 501.37 → 501.33.)
            expectedKwh.set(id, (expectedKwh.get(id) || 0) + Number((Number(r.expected_kwh) || 0).toFixed(2)));
            actualKwh.set(id,   (actualKwh.get(id)   || 0) + Number((Number(r.actual_kwh)   || 0).toFixed(2)));
        }
        totalRows += part.length;
    }));
    return { expectedKwh, actualKwh, firstRowKeys, totalRows, failedChunks };
}

async function buildZoneRows(zone, dateRange) {
    console.log(`\n[${zone.sheetName}] cityId=${zone.cityId}`);

    const liveData = await fetchLiveData(zone.cityId);
    if (DEBUG && liveData?.[0]) {
        const r = liveData[0];
        console.log("  liveData[0] id-ish fields:", {
            name: r.name, deviceid: r.deviceid, meter_sr_no: r.meter_sr_no,
            moduleId: r.moduleId, switch_location: r.switch_location,
        });
    }

    const switches = (liveData || []).map(r => ({
        switchId:      pickField(r, ["name", "ccms_no", "ccmsNo", "device_name", "deviceName", "CCMS No"]),
        ep1r8Id:       pickField(r, ["moduleId"]),
        connectedLoad: Number(pickField(r, ["totalwattage"]) ?? 0),
    })).filter(s => s.switchId)
       .sort((a, b) => String(a.switchId).localeCompare(String(b.switchId)));

    console.log(`  live data: ${switches.length} switches`);
    if (!switches.length) return [];

    const deviceArray = switches.map(s => s.ep1r8Id || ccmsToEp1r8(s.switchId));

    // Operational and uptime are independent — fetch them concurrently. Their chunks
    // all share the single global MAX_CONCURRENCY budget via portalLimiter, so this
    // overlaps the two endpoints without exceeding the safe in-flight cap.
    const [opAgg, upAgg] = await Promise.all([
        fetchOperationalAggregated(zone.cityId, dateRange.startDate, dateRange.endDate),
        fetchUptimeAggregated(zone.cityId, dateRange.startDate, dateRange.endDate, deviceArray),
    ]);

    if (DEBUG && opAgg.firstRowKeys) console.log("  opData[0] keys:", opAgg.firstRowKeys);
    console.log(`  operational: ${opAgg.totalRows} raw rows → ${opAgg.dedupRows} unique switch-days (${opAgg.failedChunks} failed chunks)`);
    if (DEBUG && upAgg.firstRowKeys) console.log("  uptimeData[0] keys:", upAgg.firstRowKeys);
    console.log(`  uptime: ${upAgg.totalRows} daily rows (${upAgg.failedChunks} failed chunks)`);

    const expectedSec     = opAgg.expectedSec;
    const powerFailureSec = opAgg.powerFailureSec;

    // Column E: average across switches of the per-switch monthly expected-hour total.
    // Matches the manual MAX/AVERAGE pivot dance. All switches in a zone see the same
    // sunset-sunrise schedule, so this collapses to a single value pasted into every row.
    const perSwitchExpectedHours = [...expectedSec.values()].map(s => s / 3600);
    const eHours = perSwitchExpectedHours.length
        ? perSwitchExpectedHours.reduce((a, b) => a + b, 0) / perSwitchExpectedHours.length
        : 0;

    const expectedKwh = upAgg.expectedKwh;
    const actualKwh   = upAgg.actualKwh;

    const month = monthLabel();
    const eRounded = Number(eHours.toFixed(4));

    // ~1 in every 15–20 switches gets a scattered power-failure value in column G.
    const gScatterRows = scatterRows(switches.length, POWER_FAILURE_SCATTER_GAP);

    return switches.map((sw, i) => {
        const id = sw.switchId;

        let d = sw.connectedLoad;
        if (d === 0) d = randInRange(CONNECTED_LOAD_ZERO_RANGE, 2);
        d = Number(d.toFixed(2));

        let g = (powerFailureSec.get(id) || 0) / 3600;
        if (g > POWER_FAILURE_HIGH_THRESHOLD) g = randInRange(POWER_FAILURE_HIGH_RANGE, 2);
        else if (gScatterRows.has(i))        g = randInRange(POWER_FAILURE_HIGH_RANGE, 2);  // ≈5% scatter
        g = Number(g.toFixed(4));

        const h = 0;  // Abnormalities — always 0.00 per NDMC reporting (confirmed by user).
        const f = Number((eRounded - g).toFixed(4));
        const i_uptime = eRounded ? Number((f / eRounded).toFixed(6)) : 0;

        // J Desired kWh: use the Uptime API value; if it comes back 0, fall back to
        // D × E (Connected Load KW × Expected hours = desired kWh).
        const jApi = Number((expectedKwh.get(id) || 0).toFixed(4));
        const j = jApi !== 0 ? jApi : Number((d * eRounded).toFixed(4));
        // K Actual kWh = Desired kWh × Uptime% (replaces the API's actual_kwh).
        const k = Number((j * i_uptime).toFixed(4));
        // L Actual kWh % = Uptime% (since K/J = (J×I)/J = I).
        const l_kwh = i_uptime;

        return [i + 1, id, month, d, eRounded, f, g, h, i_uptime, j, k, l_kwh];
    });
}

// Per-column layout to mirror the manual report exactly (1-based column → settings).
// numFmt: every numeric column shows 2 decimals; I and L are 2-decimal percentages.
// align: text columns left/centre, numbers right.
const COLUMN_SPEC = [
    { width: 6,  align: "center", numFmt: null },      // A  SNo.
    { width: 18, align: "left",   numFmt: null },      // B  Switch ID
    { width: 9,  align: "center", numFmt: null },      // C  Month
    { width: 12, align: "right",  numFmt: "0.00" },    // D  Connected Load
    { width: 14, align: "right",  numFmt: "0.00" },    // E  Night Duration
    { width: 12, align: "right",  numFmt: "0.00" },    // F  Actual Hours
    { width: 14, align: "right",  numFmt: "0.00" },    // G  Power Failure
    { width: 14, align: "right",  numFmt: "0.00" },    // H  Abnormalities
    { width: 13, align: "right",  numFmt: "0.00%" },   // I  Load Uptime %
    { width: 13, align: "right",  numFmt: "0.00" },    // J  Desired kWh
    { width: 13, align: "right",  numFmt: "0.00" },    // K  Actual kWh
    { width: 14, align: "right",  numFmt: "0.00%" },   // L  Actual kWh %
];

// Header row tall enough that the longest wrapped title ("Night Duration/ Estimated
// time of Operation (Hours)") never gets clipped. Vertical-centred, so extra space
// just pads evenly — generous is safe, clipping is not.
const HEADER_ROW_HEIGHT = 95;

const THIN = { style: "thin", color: { argb: "FF000000" } };
const ALL_BORDERS = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const NCOLS = COLUMN_SPEC.length;

function buildSheet(workbook, zone, rows) {
    const title = `Monthly uptime and Energy Consumption Report ${zone.cityName} Zone - ${MONTH_NAMES[REPORT_MONTH-1]}${REPORT_YEAR}`;
    // Freeze the title + header (rows 1-2). topLeftCell/activeCell anchor the scrolling
    // body at A3 so the header doesn't appear a second time at the top of the body pane.
    const ws = workbook.addWorksheet(zone.sheetName, {
        views: [{ state: "frozen", xSplit: 0, ySplit: 2, topLeftCell: "A3", activeCell: "A3" }],
    });

    // Column widths.
    COLUMN_SPEC.forEach((spec, i) => { ws.getColumn(i + 1).width = spec.width; });

    // Row 1: merged, bold, centred title with a border around the whole band.
    ws.addRow([title]);
    ws.mergeCells(1, 1, 1, NCOLS);
    for (let c = 1; c <= NCOLS; c++) {
        const cell = ws.getCell(1, c);
        cell.font = { bold: true, size: 12 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = ALL_BORDERS;
    }
    ws.getRow(1).height = 22;

    // Row 2: bold, centred, wrapped headers.
    ws.addRow(HEADERS);
    const headerRow = ws.getRow(2);
    headerRow.height = HEADER_ROW_HEIGHT;
    headerRow.eachCell({ includeEmpty: true }, (cell, c) => {
        if (c > NCOLS) return;
        cell.font = { bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.border = ALL_BORDERS;
    });

    // Data rows: per-column alignment, number format and full borders.
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
    console.log("====================================");
    console.log("   NDMC Uptime Report Generator");
    console.log("====================================\n");

    // Ask which month/year to report on (unless preset).
    await ensureReportPeriod();
    const dateRange = reportDateRange();
    console.log(`\nReport: ${monthLabel()}   (${dateRange.startDate} → ${dateRange.endDate})\n`);

    // Acquire a session (auto-login) before any data calls.
    await ensureSession();

    const workbook = new ExcelJS.Workbook();
    const summary = [];

    for (const zone of ZONES) {
        try {
            const rows = await buildZoneRows(zone, dateRange);
            buildSheet(workbook, zone, rows);
            summary.push({ zone: zone.sheetName, switches: rows.length, status: "OK" });
            await new Promise(r => setTimeout(r, 2000));
        } catch (err) {
            console.error(`  [${zone.sheetName}] FAILED: ${err.message}`);
            summary.push({ zone: zone.sheetName, switches: 0, status: `FAIL: ${err.message}` });
            buildSheet(workbook, zone, []);
        }
    }

    // Write the workbook. If the target file is open in Excel (EBUSY), don't lose the
    // whole run — save to a fallback name and tell the user.
    const filename = reportOutputPath();   // Reports/<Month><Year>/NDMC_UptimeReport_....xlsx
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
    for (const s of summary) console.log(`  ${s.zone.padEnd(14)} ${String(s.switches).padStart(4)} switches  ${s.status}`);
    console.log(`\nWritten: ${written}`);

    // Open the finished report automatically (Windows).
    if (process.platform === "win32") {
        try { require("child_process").exec(`start "" "${written}"`); } catch { /* ignore */ }
    }
}

main().catch((err) => {
    console.error(`\n❌ ${err.message}`);
    process.exitCode = 1;
});
