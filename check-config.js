// ============================================================================
// CONFIG CHECK — run this right after filling in config.env.
// ----------------------------------------------------------------------------
// Confirms every setting the scheduled job needs is present and readable, and
// optionally proves the portal login and the mail server actually accept the
// credentials — so a misconfiguration surfaces now rather than as a silently
// failed run on the 2nd of the month.
//
//   node check-config.js            # check settings only, no network
//   node check-config.js --live     # also test the portal login and SMTP login
//
// Nothing is emailed and no report data is downloaded in either mode.
// ============================================================================

const fs = require("fs");
const path = require("path");
const https = require("https");
const { loadConfig, maskSecret, CONFIG_FILE } = require("./load-config");

const LIVE = process.argv.includes("--live");
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

let problems = 0;
const fail = (m) => { console.log(`   ✗ ${m}`); problems++; };
const pass = (m) => console.log(`   ✓ ${m}`);
const warn = (m) => console.log(`   ! ${m}`);

console.log("=".repeat(72));
console.log("  NDMC AUTOMATION — CONFIG CHECK");
console.log("=".repeat(72));

// ---------------------------------------------------------------- file ------
console.log("\n1. config.env");
if (!fs.existsSync(CONFIG_FILE)) {
    fail(`not found at ${CONFIG_FILE}`);
    console.log("     Copy config.env.example to config.env and fill it in.");
    process.exit(1);
}
const { keys } = loadConfig({ quiet: true });
pass(`found, ${keys.length} setting(s) read`);

// ------------------------------------------------------------- portal ------
console.log("\n2. Portal credentials");
if (!process.env.NDMC_USER) fail("NDMC_USER is empty"); else pass(`NDMC_USER = ${process.env.NDMC_USER}`);
if (!process.env.NDMC_PASS) fail("NDMC_PASS is empty"); else pass(`NDMC_PASS = ${maskSecret(process.env.NDMC_PASS)}`);

// -------------------------------------------------------------- period -----
console.log("\n3. Report period");
const m = String(process.env.NDMC_MONTH || "").toLowerCase();
if (m === "last") {
    const now = new Date();
    const pm = now.getMonth();
    const p = pm === 0 ? { month: 12, year: now.getFullYear() - 1 } : { month: pm, year: now.getFullYear() };
    pass(`NDMC_MONTH = last  →  a run today would report on ${MONTH_NAMES[p.month - 1]} ${p.year}`);
} else if (Number(m) >= 1 && Number(m) <= 12) {
    warn(`NDMC_MONTH = ${m} (pinned). For the scheduled job this should be "last", or every month reports the same period.`);
} else {
    warn(`NDMC_MONTH = "${process.env.NDMC_MONTH}" — run-monthly.ps1 overrides this to "last", so it is not fatal.`);
}

// ---------------------------------------------------------------- mail -----
console.log("\n4. Email settings");
if (process.env.MAIL_ENABLED === "0") warn("MAIL_ENABLED = 0 — reports will be generated but NOT sent");
if (!process.env.MAIL_USER) fail("MAIL_USER is empty"); else pass(`MAIL_USER = ${process.env.MAIL_USER}`);
if (!process.env.MAIL_PASS) {
    fail("MAIL_PASS is empty — Gmail needs an App Password, not the account password");
} else {
    const p = process.env.MAIL_PASS.replace(/\s/g, "");
    pass(`MAIL_PASS = ${maskSecret(p)}`);
    if (p.length !== 16) warn(`a Gmail App Password is normally 16 characters — this is ${p.length}. Double-check it.`);
}
pass(`MAIL_FROM = ${process.env.MAIL_FROM || process.env.MAIL_USER || "(not set)"}`);
pass(`SMTP      = ${process.env.MAIL_HOST || "smtp.gmail.com"}:${process.env.MAIL_PORT || 587}`);

// ---------------------------------------------------------- recipients -----
console.log("\n5. Recipients");
const rf = path.join(__dirname, "recipients.txt");
if (!fs.existsSync(rf)) {
    fail("recipients.txt not found");
} else {
    const to = fs.readFileSync(rf, "utf8").split(/\r?\n/).map(s => s.trim()).filter(s => s && !s.startsWith("#"));
    if (!to.length) fail("recipients.txt has no addresses");
    else {
        pass(`${to.length} recipient(s):`);
        for (const t of to) {
            const ok = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t);
            console.log(`        ${ok ? "•" : "✗ INVALID:"} ${t}`);
            if (!ok) problems++;
        }
    }
}

// -------------------------------------------------------------- live -------
(async () => {
    if (LIVE && !problems) {
        console.log("\n6. Live checks");

        // Portal login only — one small request, no report data downloaded.
        try {
            const axios = require("axios");
            const BASE = "https://smartlight.citilight.co:446";
            const agent = new https.Agent({ keepAlive: false });
            const common = { httpsAgent: agent, timeout: 60000, maxRedirects: 0, validateStatus: () => true };
            const seed = await axios.get(`${BASE}/smartlight/login`, common);
            const grab = (sc) => {
                for (const c of (Array.isArray(sc) ? sc : [sc]).filter(Boolean)) {
                    const mm = /JSESSIONID=([^;]+)/i.exec(c);
                    if (mm) return mm[1];
                }
                return null;
            };
            let sid = grab(seed.headers["set-cookie"]);
            const form = `username=${encodeURIComponent(process.env.NDMC_USER)}&password=${encodeURIComponent(process.env.NDMC_PASS)}`;
            const res = await axios.post(`${BASE}/smartlight/login`, form, {
                ...common,
                headers: {
                    "User-Agent": "Mozilla/5.0",
                    "Content-Type": "application/x-www-form-urlencoded",
                    Origin: BASE, Referer: `${BASE}/smartlight/login`,
                    ...(sid ? { Cookie: `JSESSIONID=${sid}` } : {}),
                },
            });
            sid = grab(res.headers["set-cookie"]) || sid;
            if (!sid) fail("portal login returned no session cookie");
            else {
                // One tiny data call proves the session is genuinely authenticated.
                const probe = await axios.post(`${BASE}/smartlight/getListViewData_v1`,
                    { cityId: "2", deviceType: "1", zoneName: "0", wardName: "0", streetName: "0", userId: "10" },
                    { httpsAgent: agent, timeout: 60000, headers: {
                        "User-Agent": "Mozilla/5.0", "Content-Type": "application/json",
                        "X-Requested-With": "XMLHttpRequest", Origin: BASE,
                        Referer: `${BASE}/smartlight/livedatafeed`, Cookie: `JSESSIONID=${sid}`,
                    } });
                if (Array.isArray(probe.data)) pass(`portal login OK (SP zone returned ${probe.data.length} switches)`);
                else fail("portal login failed — check NDMC_USER / NDMC_PASS");
            }
        } catch (err) {
            fail(`portal check failed: ${err.code || ""} ${err.message}`.trim());
        }

        // SMTP login only — verify() authenticates without sending anything.
        try {
            const nodemailer = require("nodemailer");
            const port = Number(process.env.MAIL_PORT) || 587;
            const t = nodemailer.createTransport({
                host: process.env.MAIL_HOST || "smtp.gmail.com",
                port, secure: port === 465,
                auth: { user: process.env.MAIL_USER, pass: String(process.env.MAIL_PASS || "").replace(/\s/g, "") },
            });
            await t.verify();
            pass("SMTP login OK (nothing was sent)");
        } catch (err) {
            fail(`SMTP login failed: ${err.message}`);
            console.log("     Gmail rejects normal passwords — you need an App Password.");
        }
    } else if (LIVE) {
        console.log("\n6. Live checks — skipped, fix the problems above first.");
    }

    console.log("\n" + "=".repeat(72));
    if (problems) {
        console.log(`  ${problems} PROBLEM(S) — the scheduled run would fail.`);
        process.exit(1);
    }
    console.log("  ALL CHECKS PASSED" + (LIVE ? " — portal and mail credentials both work." : ""));
    if (!LIVE) console.log("  Now prove the credentials work:   node check-config.js --live");
    console.log("=".repeat(72));
})();
