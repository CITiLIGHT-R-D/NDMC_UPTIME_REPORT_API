// ============================================================================
// CONFIG LOADER
// ----------------------------------------------------------------------------
// Reads config.env into process.env so the Node scripts work when run BY HAND,
// not only when launched from run-monthly.ps1 (which loads the same file
// itself). Real environment variables always win, so a value set on the command
// line still overrides the file.
//
// Format: KEY=value, one per line. "#" starts a comment. Surrounding quotes are
// stripped. Everything after the FIRST "=" is the value, so passwords may
// contain "=" safely.
// ============================================================================

const fs = require("fs");
const path = require("path");

const CONFIG_FILE = path.join(__dirname, "config.env");

function loadConfig({ quiet = false } = {}) {
    if (!fs.existsSync(CONFIG_FILE)) {
        if (!quiet) console.log("(no config.env found — using environment variables only)");
        return { loaded: false, keys: [] };
    }
    const keys = [];
    for (const raw of fs.readFileSync(CONFIG_FILE, "utf8").split(/\r?\n/)) {
        const line = raw.trim();
        if (!line || line.startsWith("#")) continue;
        const eq = line.indexOf("=");
        if (eq <= 0) continue;
        const key = line.slice(0, eq).trim();
        let val = line.slice(eq + 1).trim();
        if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
            val = val.slice(1, -1);
        }
        if (process.env[key] === undefined || process.env[key] === "") {
            process.env[key] = val;
        }
        keys.push(key);
    }
    return { loaded: true, keys };
}

// Never print a secret in full — enough to confirm it is the right value, no more.
function maskSecret(v) {
    if (!v) return "(not set)";
    if (v.length <= 4) return "****";
    return `${v.slice(0, 2)}${"*".repeat(Math.max(4, v.length - 4))}${v.slice(-2)}  (${v.length} chars)`;
}

module.exports = { loadConfig, maskSecret, CONFIG_FILE };
