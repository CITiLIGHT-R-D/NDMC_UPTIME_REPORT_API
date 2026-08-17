// ============================================================================
// MONTHLY SCHEDULER (node-cron)
// ----------------------------------------------------------------------------
// Long-running process that fires the monthly report run. It does NOT reimplement
// the run itself — it launches run-monthly.ps1, which stays the single place that
// generates, verifies and emails the reports.
//
// WHY NOT node-cron ALONE:
//   node-cron only fires while this process is alive. On a PC that gets shut
//   down, a monthly job would simply be lost — and a monthly job only gets 12
//   chances a year. So this adds a CATCH-UP pass on startup.
//
// WHY THIS BEATS Task Scheduler ALONE:
//   Task Scheduler knows only whether the trigger FIRED. This checks whether the
//   report actually EXISTS and SUCCEEDED, so a run that fired and failed is
//   retried at the next opportunity instead of being written off.
//
// The pairing: one Task Scheduler entry with an "At startup" trigger keeps this
// process alive across reboots; this process owns the schedule and the catch-up.
//
// Usage:
//   node scheduler.js                 # run in the foreground (Ctrl+C to stop)
//   node scheduler.js --run-now       # run the report immediately, then exit
//   node scheduler.js --status        # show schedule + history, then exit
//   node scheduler.js --no-catchup    # start without the startup catch-up pass
//
// Schedule and behaviour are configurable in config.env:
//   SCHEDULE_CRON=0 1 2 * *     (minute hour day month weekday) default: 01:00 on the 2nd
//   SCHEDULE_CATCHUP=1          (0 disables the startup catch-up)
// ============================================================================

const cron = require("node-cron");
const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");

require("./load-config").loadConfig({ quiet: true });

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const FULL_MONTH_NAMES = ["January","February","March","April","May","June",
                          "July","August","September","October","November","December"];

// 01:00 on the 2nd of every month. The 2nd (not the 1st) because nights span
// midnight — the final night of the month is not complete until the 1st morning.
const CRON_EXPR   = process.env.SCHEDULE_CRON || "0 1 2 * *";
const DO_CATCHUP  = process.env.SCHEDULE_CATCHUP !== "0" && !process.argv.includes("--no-catchup");
const STATE_FILE  = path.join(__dirname, ".scheduler-state.json");
const WRAPPER     = path.join(__dirname, "run-monthly.ps1");

// --------------------------------------------------------------- logging ----
const logDir = path.join(__dirname, "Logs");
function log(msg, level = "INFO") {
    const line = `${new Date().toISOString().replace("T", " ").slice(0, 19)} [SCHED/${level}] ${msg}`;
    console.log(line);
    try {
        fs.mkdirSync(logDir, { recursive: true });
        fs.appendFileSync(path.join(logDir, "scheduler.log"), line + "\n", "utf8");
    } catch { /* logging must never take the scheduler down */ }
}

// ----------------------------------------------------------------- state ----
function readState() {
    try { return JSON.parse(fs.readFileSync(STATE_FILE, "utf8")); }
    catch { return { completed: {} }; }
}
function writeState(s) {
    try { fs.writeFileSync(STATE_FILE, JSON.stringify(s, null, 2), "utf8"); }
    catch (e) { log(`could not write state: ${e.message}`, "WARN"); }
}

// The month a run should cover: the one that just ended.
function targetPeriod(now = new Date()) {
    const m = now.getMonth();               // 0-11 → already the previous month in 1-12
    return m === 0
        ? { month: 12, year: now.getFullYear() - 1 }
        : { month: m,  year: now.getFullYear() };
}
const periodKey   = (p) => `${p.year}-${String(p.month).padStart(2, "0")}`;
const periodLabel = (p) => `${FULL_MONTH_NAMES[p.month - 1]} ${p.year}`;

// Ground truth, not just bookkeeping: a period counts as done only if BOTH report
// files are actually on disk. State alone could be stale or hand-edited.
function reportsExist(p) {
    const dir = path.join(__dirname, "Reports", `${MONTH_NAMES[p.month - 1]}${p.year}`);
    const files = [
        path.join(dir, `NDMC_UptimeReport_${MONTH_NAMES[p.month - 1]}${p.year}.xlsx`),
        path.join(dir, `Operational Hour Report ${FULL_MONTH_NAMES[p.month - 1]} ${p.year}.xlsx`),
    ];
    const present = files.filter(f => fs.existsSync(f));
    return { all: present.length === files.length, present: present.length, total: files.length };
}

// Files on disk are the ground truth — if both reports are already there, there is
// nothing to do, even when this scheduler has no record of producing them (someone
// may have run them by hand). The one exception: a recorded FAILURE means the files
// that exist may be partial, so that period is retried.
function isDone(p) {
    const st = readState();
    const rec = st.completed[periodKey(p)];
    const files = reportsExist(p);
    const failedBefore = Boolean(rec && rec.status === "failed");
    return { done: files.all && !failedBefore, rec, files };
}

// ------------------------------------------------------------------- run ----
let running = false;

function runReport(reason) {
    if (running) { log(`already running — ignoring trigger (${reason})`, "WARN"); return; }

    const p = targetPeriod();
    running = true;
    log(`starting run for ${periodLabel(p)} (${reason})`);

    // PowerShell wrapper does the real work: generate, verify, email.
    const child = spawn("powershell.exe",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", WRAPPER],
        { cwd: __dirname, windowsHide: true });

    child.stdout.on("data", d => process.stdout.write(d));
    child.stderr.on("data", d => process.stderr.write(d));

    child.on("close", (code) => {
        running = false;
        const files = reportsExist(p);
        const ok = code === 0 && files.all;
        const st = readState();
        st.completed[periodKey(p)] = {
            status: ok ? "success" : "failed",
            exitCode: code,
            filesPresent: `${files.present}/${files.total}`,
            at: new Date().toISOString(),
            attempts: ((st.completed[periodKey(p)] || {}).attempts || 0) + 1,
        };
        writeState(st);
        log(ok
            ? `run for ${periodLabel(p)} SUCCEEDED (both reports present)`
            : `run for ${periodLabel(p)} FAILED (exit ${code}, ${files.present}/${files.total} reports present) — will retry at the next startup or schedule`,
            ok ? "INFO" : "ERROR");
    });

    child.on("error", (err) => {
        running = false;
        log(`could not launch run-monthly.ps1: ${err.message}`, "ERROR");
    });
}

// -------------------------------------------------------------- catch-up ----
// Runs at startup. Covers the cases node-cron alone cannot: the PC was off when
// the schedule fired, the process was not running, or a previous run failed.
function catchUp() {
    const p = targetPeriod();
    const now = new Date();
    const { done, rec, files } = isDone(p);

    if (done) {
        log(`catch-up: ${periodLabel(p)} already complete — nothing to do`);
        return;
    }
    // Before the scheduled day, the month is not yet due; wait for the trigger.
    const scheduledDay = Number((CRON_EXPR.split(/\s+/)[2] || "2").replace(/\D/g, "")) || 2;
    if (now.getDate() < scheduledDay) {
        log(`catch-up: ${periodLabel(p)} not due yet (runs on day ${scheduledDay}, today is the ${now.getDate()})`);
        return;
    }
    const why = rec && rec.status === "failed"
        ? `previous attempt failed (exit ${rec.exitCode}, ${rec.filesPresent} reports)`
        : files.present > 0
            ? `only ${files.present}/${files.total} reports present`
            : "no run recorded for this period";
    log(`catch-up: ${periodLabel(p)} is overdue — ${why}`, "WARN");
    runReport("startup catch-up");
}

// ---------------------------------------------------------------- status ----
function showStatus() {
    const p = targetPeriod();
    const { done, rec, files } = isDone(p);
    const st = readState();
    console.log("=".repeat(72));
    console.log("  NDMC SCHEDULER STATUS");
    console.log("=".repeat(72));
    console.log(`  Schedule       : ${CRON_EXPR}   (${cron.validate(CRON_EXPR) ? "valid" : "INVALID"})`);
    console.log(`  Meaning        : 01:00 on day ${(CRON_EXPR.split(/\s+/)[2] || "2")} of every month`);
    console.log(`  Startup catchup: ${DO_CATCHUP ? "enabled" : "disabled"}`);
    console.log(`  Current target : ${periodLabel(p)}`);
    console.log(`  Reports on disk: ${files.present}/${files.total}`);
    console.log(`  Status         : ${done ? "COMPLETE" : "PENDING"}`);
    if (rec) console.log(`  Last attempt   : ${rec.status} at ${rec.at} (attempt ${rec.attempts})`);
    const hist = Object.keys(st.completed).sort().reverse().slice(0, 12);
    if (hist.length) {
        console.log("\n  History:");
        for (const k of hist) {
            const r = st.completed[k];
            console.log(`    ${k}  ${String(r.status).padEnd(8)} ${r.filesPresent} reports  ${r.at}`);
        }
    }
    console.log("=".repeat(72));
}

// ------------------------------------------------------------------ main ----
if (process.argv.includes("--status")) { showStatus(); process.exit(0); }

if (!cron.validate(CRON_EXPR)) {
    log(`SCHEDULE_CRON is not a valid cron expression: "${CRON_EXPR}"`, "ERROR");
    process.exit(1);
}
if (!fs.existsSync(WRAPPER)) {
    log(`run-monthly.ps1 not found at ${WRAPPER}`, "ERROR");
    process.exit(1);
}

if (process.argv.includes("--run-now")) {
    log("manual run requested (--run-now)");
    runReport("manual --run-now");
    // Keep the process alive until the child finishes, then exit with its result.
    const wait = setInterval(() => { if (!running) { clearInterval(wait); process.exit(0); } }, 1000);
} else {
    log("=".repeat(60));
    log("NDMC monthly scheduler started");
    log(`schedule: ${CRON_EXPR}  (01:00 on day ${(CRON_EXPR.split(/\s+/)[2] || "2")} of each month)`);
    log(`catch-up on startup: ${DO_CATCHUP ? "enabled" : "disabled"}`);
    log("=".repeat(60));

    cron.schedule(CRON_EXPR, () => runReport("scheduled trigger"));

    if (DO_CATCHUP) catchUp();

    // A monthly job is idle almost all the time; a periodic line proves it is
    // still alive rather than silently dead.
    setInterval(() => {
        const p = targetPeriod();
        log(`alive — target ${periodLabel(p)}, ${isDone(p).done ? "complete" : "pending"}`);
    }, 12 * 60 * 60 * 1000);

    const bye = (sig) => { log(`received ${sig} — shutting down`); process.exit(0); };
    process.on("SIGINT",  () => bye("SIGINT"));
    process.on("SIGTERM", () => bye("SIGTERM"));
}
