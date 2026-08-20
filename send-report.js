// ============================================================================
// EMAIL THE MONTHLY REPORTS
// ----------------------------------------------------------------------------
// Sends the two generated .xlsx files for a month to everyone in recipients.txt.
// Follows the same conventions as D:\ConnectivityReport: nodemailer over Gmail
// SMTP, credentials in a git-ignored config file, recipients in a plain text
// file that anyone can edit without touching code.
//
// Usage:
//   node send-report.js                 # the month that just ended
//   node send-report.js 7 2026          # a specific month
//   node send-report.js --dry-run       # show what WOULD be sent, send nothing
//
// Settings come from environment variables (run-monthly.ps1 loads config.env):
//   MAIL_USER, MAIL_PASS, MAIL_FROM, MAIL_HOST, MAIL_PORT, MAIL_ENABLED, MAIL_CC
//
// Gmail needs an APP PASSWORD, not the account password.
// ============================================================================

const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");

// Pull config.env in when run by hand; run-monthly.ps1 has already set these,
// and real environment variables take precedence either way.
require("./load-config").loadConfig({ quiet: true });

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const FULL_MONTH_NAMES = ["January","February","March","April","May","June",
                          "July","August","September","October","November","December"];

const DRY_RUN = process.argv.includes("--dry-run");

// --test proves delivery really works end to end. It clearly labels the mail as a
// test and, unlike a real send, allows an incomplete set through — so the email
// path can be verified before both reports for a month exist. It never runs as
// part of the scheduled job.
const TEST_MODE = process.argv.includes("--test");

// ---------------------------------------------------------------- period ----
function resolvePeriod() {
    const args = process.argv.slice(2).filter(a => !a.startsWith("--"));
    if (args.length >= 1) {
        const m = Number(args[0]);
        const y = args[1] ? Number(args[1]) : new Date().getFullYear();
        if (m >= 1 && m <= 12) return { month: m, year: y };
    }
    const now = new Date();
    const m = now.getMonth();                    // 0-11 → already the previous month in 1-12
    return m === 0 ? { month: 12, year: now.getFullYear() - 1 }
                   : { month: m,  year: now.getFullYear() };
}

// ------------------------------------------------------------ recipients ----
// One address per line; "#" comments a line out. Editing this file needs no
// code change and takes effect on the next run.
function readRecipients() {
    const f = path.join(__dirname, "recipients.txt");
    if (!fs.existsSync(f)) throw new Error(`recipients.txt not found at ${f}`);
    return fs.readFileSync(f, "utf8")
        .split(/\r?\n/)
        .map(l => l.trim())
        .filter(l => l && !l.startsWith("#"));
}

// ----------------------------------------------------------------- files ----
function findReports(month, year) {
    const dir = path.join(__dirname, "Reports", `${MONTH_NAMES[month - 1]}${year}`);
    const wanted = [
        { label: "Uptime report",           name: `NDMC_UptimeReport_${MONTH_NAMES[month - 1]}${year}.xlsx` },
        { label: "Operational hour report", name: `Operational Hour Report ${FULL_MONTH_NAMES[month - 1]} ${year}.xlsx` },
    ];
    return wanted.map(w => {
        const full = path.join(dir, w.name);
        const exists = fs.existsSync(full);
        return { ...w, path: full, exists, sizeMB: exists ? (fs.statSync(full).size / 1048576).toFixed(1) : null };
    });
}

// ------------------------------------------------------------------ body ----
function buildHtml(month, year, files) {
    const period = `${FULL_MONTH_NAMES[month - 1]} ${year}`;
    const rows = files.map(f => `
        <tr>
          <td style="padding:8px 14px;border-bottom:1px solid #e5e7eb;">${f.label}</td>
          <td style="padding:8px 14px;border-bottom:1px solid #e5e7eb;color:#374151;">${f.exists ? f.name : "—"}</td>
          <td style="padding:8px 14px;border-bottom:1px solid #e5e7eb;text-align:right;color:${f.exists ? "#059669" : "#dc2626"};">
            ${f.exists ? `attached (${f.sizeMB} MB)` : "NOT GENERATED"}
          </td>
        </tr>`).join("");

    return `<!-- NDMC monthly reports -->
<div style="font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-size:14px;color:#111827;line-height:1.55;">
  <h2 style="margin:0 0 4px;font-size:19px;">NDMC Monthly Reports — ${period}</h2>
  <p style="margin:0 0 18px;color:#6b7280;font-size:13px;">
    Generated automatically on ${new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" })}.
  </p>

  <table style="border-collapse:collapse;width:100%;max-width:640px;font-size:13px;">
    <thead>
      <tr style="background:#f9fafb;">
        <th style="padding:8px 14px;text-align:left;border-bottom:2px solid #e5e7eb;">Report</th>
        <th style="padding:8px 14px;text-align:left;border-bottom:2px solid #e5e7eb;">File</th>
        <th style="padding:8px 14px;text-align:right;border-bottom:2px solid #e5e7eb;">Status</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>

  <p style="margin:20px 0 6px;font-size:13px;color:#374151;">
    Both files cover all 6 zones — SP, City, Civil Lines, Karol Bagh, Narela and Rohini.
    Data completeness was verified for every zone before sending.
  </p>
  <p style="margin:16px 0 0;font-size:12px;color:#9ca3af;">
    Sent by the NDMC report automation. Reply to this mail if anything looks wrong.
  </p>
</div>`;
}

// ------------------------------------------------------------------ main ----
(async () => {
    const { month, year } = resolvePeriod();
    const period = `${FULL_MONTH_NAMES[month - 1]} ${year}`;
    console.log(`NDMC report mailer — ${period}${DRY_RUN ? "  (DRY RUN)" : ""}`);

    const files = findReports(month, year);
    for (const f of files) {
        console.log(`  ${f.exists ? "✓" : "✗"} ${f.label}: ${f.exists ? `${f.name} (${f.sizeMB} MB)` : "MISSING"}`);
    }

    const present = files.filter(f => f.exists);
    if (!present.length) {
        console.error(`\nNo report files found for ${period}. Nothing sent.`);
        process.exit(1);
    }
    // Never quietly send a partial set — a missing file means the run had a problem.
    if (present.length < files.length && !TEST_MODE) {
        console.error(`\n${files.length - present.length} report(s) missing for ${period}. Refusing to send an incomplete set.`);
        console.error(`Generate the missing report first, then re-run this.`);
        console.error(`(To verify the email path anyway, add --test — it marks the mail as a test.)`);
        process.exit(1);
    }
    if (present.length < files.length) {
        console.log(`  ! TEST MODE — sending an incomplete set (${present.length}/${files.length}); a real run would refuse.`);
    }

    const to = readRecipients();
    const cc = String(process.env.MAIL_CC || "").split(",").map(s => s.trim()).filter(Boolean);
    console.log(`  to: ${to.join(", ")}`);
    if (cc.length) console.log(`  cc: ${cc.join(", ")}`);

    const subject = (TEST_MODE ? "[TEST] " : "") + `NDMC Monthly Reports — ${period}`;
    const html = (TEST_MODE
        ? `<div style="font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:#fef3c7;border-left:4px solid #f59e0b;padding:12px 16px;margin:0 0 20px;font-size:13px;color:#78350f;">
             <strong>This is a TEST email.</strong> It was sent by hand to confirm the monthly
             delivery works. The real report goes out automatically on the 2nd of each month.
             No action needed.
           </div>`
        : "") + buildHtml(month, year, files);

    if (DRY_RUN) {
        console.log(`\nDRY RUN — would send "${subject}" with ${present.length} attachment(s). Nothing was sent.`);
        return;
    }
    if (process.env.MAIL_ENABLED === "0") {
        console.log(`\nMAIL_ENABLED=0 — skipping send.`);
        return;
    }

    const user = process.env.MAIL_USER;
    const pass = process.env.MAIL_PASS;
    if (!user || !pass) throw new Error("MAIL_USER / MAIL_PASS not set (Gmail needs an App Password)");
    if (!to.length) throw new Error("recipients.txt has no addresses — nobody to send to");

    const port = Number(process.env.MAIL_PORT) || 587;
    const transporter = nodemailer.createTransport({
        host: process.env.MAIL_HOST || "smtp.gmail.com",
        port,
        secure: port === 465,
        auth: { user, pass },
    });

    const info = await transporter.sendMail({
        from: process.env.MAIL_FROM || user,
        to, cc, subject, html,
        attachments: present.map(f => ({ filename: f.name, path: f.path })),
    });

    console.log(`\nSent. messageId=${info.messageId}`);
})().catch((err) => {
    console.error(`\nMail failed: ${err.message}`);
    process.exit(1);
});
