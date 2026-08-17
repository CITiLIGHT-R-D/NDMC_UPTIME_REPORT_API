// ============================================================================
// FAILURE ALERT
// ----------------------------------------------------------------------------
// Emails the recipients when the scheduled monthly run did NOT produce a
// complete set of reports. A monthly job that quietly stops working is worse
// than no job at all — this is what makes a failure visible.
//
// Called by run-monthly.ps1 with the failure reasons as arguments:
//   node send-alert.js "Uptime report incomplete after all attempts"
// ============================================================================

const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");

require("./load-config").loadConfig({ quiet: true });

const FULL_MONTH_NAMES = ["January","February","March","April","May","June",
                          "July","August","September","October","November","December"];

function previousMonth() {
    const now = new Date();
    const m = now.getMonth();
    return m === 0 ? { month: 12, year: now.getFullYear() - 1 }
                   : { month: m,  year: now.getFullYear() };
}

function readRecipients() {
    const f = path.join(__dirname, "recipients.txt");
    if (!fs.existsSync(f)) return [];
    return fs.readFileSync(f, "utf8")
        .split(/\r?\n/).map(l => l.trim()).filter(l => l && !l.startsWith("#"));
}

// Newest log file, so the alert can quote the tail without anyone opening the box.
function latestLogTail(lines = 25) {
    const dir = path.join(__dirname, "Logs");
    if (!fs.existsSync(dir)) return null;
    const logs = fs.readdirSync(dir)
        .filter(f => f.endsWith(".log"))
        .map(f => ({ f, m: fs.statSync(path.join(dir, f)).mtimeMs }))
        .sort((a, b) => b.m - a.m);
    if (!logs.length) return null;
    const content = fs.readFileSync(path.join(dir, logs[0].f), "utf8").split(/\r?\n/);
    return { name: logs[0].f, tail: content.slice(-lines).join("\n") };
}

(async () => {
    const reasons = process.argv.slice(2).filter(Boolean);
    const { month, year } = previousMonth();
    const period = `${FULL_MONTH_NAMES[month - 1]} ${year}`;

    const to = readRecipients();
    const user = process.env.MAIL_USER, pass = process.env.MAIL_PASS;
    if (!user || !pass || !to.length || process.env.MAIL_ENABLED === "0") {
        console.log("Alert not sent (mail not configured or disabled).");
        return;
    }

    const log = latestLogTail();
    const html = `
<div style="font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-size:14px;color:#111827;line-height:1.55;">
  <h2 style="margin:0 0 4px;font-size:19px;color:#b91c1c;">⚠ NDMC report run FAILED — ${period}</h2>
  <p style="margin:0 0 16px;color:#6b7280;font-size:13px;">
    The scheduled run on ${new Date().toLocaleString("en-GB")} did not produce a complete set.
    <strong>No reports were sent.</strong>
  </p>

  <p style="margin:0 0 6px;font-weight:600;">What went wrong</p>
  <ul style="margin:0 0 18px;padding-left:20px;color:#374151;">
    ${reasons.map(r => `<li>${r}</li>`).join("") || "<li>Unspecified failure</li>"}
  </ul>

  <p style="margin:0 0 6px;font-weight:600;">What to do</p>
  <ol style="margin:0 0 18px;padding-left:20px;color:#374151;">
    <li>Re-run manually: <code>.\\run-monthly.ps1</code> — completed downloads are cached, so a retry is quick.</li>
    <li>Verify with <code>node check-completeness.js</code> — all 6 zones must read OK.</li>
    <li>If zones keep failing, the portal is likely under load. Try again later.</li>
  </ol>

  ${log ? `<p style="margin:0 0 6px;font-weight:600;">Last lines of ${log.name}</p>
  <pre style="background:#f3f4f6;padding:12px;border-radius:6px;font-size:11px;overflow-x:auto;color:#374151;">${
      log.tail.replace(/[<>&]/g, c => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[c]))
  }</pre>` : ""}
</div>`;

    const port = Number(process.env.MAIL_PORT) || 587;
    const transporter = nodemailer.createTransport({
        host: process.env.MAIL_HOST || "smtp.gmail.com",
        port, secure: port === 465,
        auth: { user, pass },
    });

    await transporter.sendMail({
        from: process.env.MAIL_FROM || user,
        to,
        subject: `⚠ NDMC report run FAILED — ${period}`,
        html,
    });
    console.log(`Failure alert sent to ${to.join(", ")}`);
})().catch((err) => {
    console.error(`Alert failed: ${err.message}`);
    process.exit(1);
});
