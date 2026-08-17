<#
================================================================================
  NDMC MONTHLY REPORTS — unattended runner
================================================================================
  Called by Windows Task Scheduler on the 2nd of each month. Generates BOTH
  reports for the month that just ended, verifies them, and retries until they
  are complete.

  Why it retries: the portal drops requests under load. Every chunk that does
  download is checkpointed to .cache\, so a retry only fetches what is still
  missing — usually a couple of minutes rather than a full re-run. Retrying is
  therefore cheap, and is what turns an unreliable download into a reliable job.

  Setup (one time):
    1. Copy config.env.example to config.env and fill in the portal credentials.
    2. Lock it down so only the service account can read it:
         icacls config.env /inheritance:r /grant:r "$env:USERNAME:(R)"
    3. Register the scheduled task — see KT\AUTOMATION.md.

  Manual use:
    .\run-monthly.ps1                 # previous month
    .\run-monthly.ps1 -Month 7 -Year 2026
================================================================================
#>

[CmdletBinding()]
param(
    [int]$Month = 0,                 # 0 = the month that just ended
    [int]$Year  = 0,
    [int]$MaxAttempts = 4,           # per report; each retry resumes from .cache
    [switch]$SkipOperational
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

# ---------------------------------------------------------------- logging ----
$logDir = Join-Path $PSScriptRoot 'Logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$stamp  = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$logFile = Join-Path $logDir "monthly_$stamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding utf8
}

# ------------------------------------------------------------------- lock ----
# One run at a time, always. Two concurrent runs would double the load on a
# portal that already fails at three simultaneous requests.
$lockFile = Join-Path $PSScriptRoot '.run.lock'
if (Test-Path $lockFile) {
    $age = (Get-Date) - (Get-Item $lockFile).LastWriteTime
    if ($age.TotalHours -lt 8) {
        Write-Log "Another run started $([int]$age.TotalMinutes) min ago (.run.lock). Exiting." 'WARN'
        exit 0
    }
    Write-Log "Stale lock ($([int]$age.TotalHours)h old) — taking over." 'WARN'
    Remove-Item $lockFile -Force
}
Set-Content -Path $lockFile -Value $PID -Encoding utf8

try {
    Write-Log "=============================================================="
    Write-Log " NDMC monthly report run"
    Write-Log "=============================================================="

    # --------------------------------------------------------- credentials ---
    $cfg = Join-Path $PSScriptRoot 'config.env'
    if (-not (Test-Path $cfg)) {
        Write-Log "config.env not found. Copy config.env.example and fill it in." 'ERROR'
        exit 1
    }
    Get-Content $cfg | ForEach-Object {
        $line = $_.Trim()
        if ($line -and -not $line.StartsWith('#')) {
            $i = $line.IndexOf('=')
            if ($i -gt 0) {
                $k = $line.Substring(0, $i).Trim()
                $v = $line.Substring($i + 1).Trim().Trim('"')
                Set-Item -Path "env:$k" -Value $v
            }
        }
    }
    if (-not $env:NDMC_USER -or -not $env:NDMC_PASS) {
        Write-Log "NDMC_USER / NDMC_PASS missing from config.env." 'ERROR'
        exit 1
    }

    # -------------------------------------------------------------- period ---
    # Default: the month that just ended. The Node scripts work this out
    # themselves from NDMC_MONTH=last, so January rolls back to December safely.
    if ($Month -ge 1 -and $Month -le 12) {
        $env:NDMC_MONTH = "$Month"
        if ($Year -ge 2000) { $env:NDMC_YEAR = "$Year" } else { Remove-Item env:NDMC_YEAR -ErrorAction SilentlyContinue }
        Write-Log "Reporting on month $Month $(if ($Year -ge 2000) { $Year } else { '(current year)' }) (explicit)"
    } else {
        $env:NDMC_MONTH = 'last'
        Remove-Item env:NDMC_YEAR -ErrorAction SilentlyContinue
        Write-Log "Reporting on the month that just ended"
    }
    $env:NDMC_UNATTENDED = '1'

    # ---------------------------------------------------------------- node ---
    $node = (Get-Command node -ErrorAction SilentlyContinue)
    if (-not $node) { Write-Log "Node.js not found on PATH." 'ERROR'; exit 1 }
    if (-not (Test-Path (Join-Path $PSScriptRoot 'node_modules\exceljs'))) {
        Write-Log "Installing dependencies (first run)..."
        & npm install 2>&1 | ForEach-Object { Write-Log $_ 'NPM' }
    }

    # Run a Node script, streaming its output into the log. Returns the exit code.
    function Invoke-Report {
        param([string]$Script)
        & node $Script 2>&1 | ForEach-Object { Write-Log $_ 'NODE' }
        return $LASTEXITCODE
    }

    $failures = @()

    # ------------------------------------------------------ uptime report ----
    Write-Log ""
    Write-Log "--- Uptime report ---"
    $ok = $false
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log "Attempt $attempt of $MaxAttempts"
        $code = Invoke-Report 'NdmcUptimeReport.js'
        if ($code -eq 0) {
            # The generator exits 0 only when every zone completed; verify anyway.
            & node check-completeness.js 2>&1 | ForEach-Object { Write-Log $_ 'CHECK' }
            if ($LASTEXITCODE -eq 0) { $ok = $true; Write-Log "Uptime report complete and verified."; break }
            Write-Log "Completeness check failed — retrying (cached chunks are reused)." 'WARN'
        } else {
            Write-Log "Generator reported incomplete zones — retrying (cached chunks are reused)." 'WARN'
        }
        if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds 120 }
    }
    if (-not $ok) { $failures += 'Uptime report incomplete after all attempts' }

    # ------------------------------------------------ operational report ----
    if (-not $SkipOperational) {
        Write-Log ""
        Write-Log "--- Operational hour report ---"
        $opOk = $false
        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            Write-Log "Attempt $attempt of $MaxAttempts"
            if ((Invoke-Report 'NdmcOperationalReport.js') -eq 0) { $opOk = $true; Write-Log "Operational report complete."; break }
            Write-Log "Operational report failed — retrying." 'WARN'
            if ($attempt -lt $MaxAttempts) { Start-Sleep -Seconds 120 }
        }
        if (-not $opOk) { $failures += 'Operational report failed after all attempts' }
    }

    # ------------------------------------------------------------- outcome ---
    # -------------------------------------------------------------- email ---
    # Only ever mail a verified, complete set. A failed run must not quietly
    # deliver partial data — that is the whole problem this pipeline exists to
    # prevent — so on failure we alert instead of attaching anything.
    Write-Log ""
    if ($failures.Count -eq 0) {
        Write-Log "--- Emailing reports ---"
        & node send-report.js 2>&1 | ForEach-Object { Write-Log $_ 'MAIL' }
        if ($LASTEXITCODE -ne 0) { $failures += 'Reports generated but the email failed to send' }
    } else {
        Write-Log "Skipping email — the run did not produce a complete set." 'WARN'
    }

    # ------------------------------------------------------------- outcome ---
    Write-Log ""
    Write-Log "=============================================================="
    if ($failures.Count -eq 0) {
        Write-Log "SUCCESS — both reports generated, verified and emailed."
        Write-Log "Output: $(Join-Path $PSScriptRoot 'Reports')"
        Write-Log "=============================================================="
        exit 0
    }

    foreach ($f in $failures) { Write-Log $f 'ERROR' }
    Write-Log "RUN FAILED — reports were NOT sent. Log: $logFile" 'ERROR'
    Write-Log "=============================================================="

    # Tell a human it broke, otherwise a silently-dead monthly job is worse than none.
    try {
        & node send-alert.js @($failures) 2>&1 | ForEach-Object { Write-Log $_ 'ALERT' }
    } catch {
        Write-Log "Could not send the failure alert: $_" 'WARN'
    }
    exit 1
}
finally {
    Remove-Item $lockFile -Force -ErrorAction SilentlyContinue
}
