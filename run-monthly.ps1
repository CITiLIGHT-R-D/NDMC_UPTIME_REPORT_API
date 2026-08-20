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

# ------------------------------------------------------- keep-awake ----------
# A full run takes hours. WakeToRun only wakes the machine to START the job - it
# does nothing to stop Windows sleeping DURING it, which is exactly how the
# 17 Aug 2026 test died: the PC slept at ~20:20 and every request then failed
# with ENOTFOUND (DNS gone), part-way through the third zone.
#
# SetThreadExecutionState with ES_CONTINUOUS | ES_SYSTEM_REQUIRED tells Windows
# this thread needs the system kept alive. The display is deliberately NOT held
# on - the screen can still blank, the machine just must not sleep. Cleared
# again in the finally block.
#
# NOT ENOUGH ON ITS OWN: this machine uses Modern Standby (S0ix), which
# SetThreadExecutionState does NOT prevent - confirmed by Kernel-Power event 506
# ("entering Modern Standby") landing exactly when two test runs died. So the
# idle standby timeout is also set to "never" below, and restored afterwards.
$keepAwake = $null
try {
    $sig = @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern uint SetThreadExecutionState(uint esFlags);
'@
    $keepAwake = Add-Type -MemberDefinition $sig -Name 'PowerMgmt' -Namespace 'Win32Native' -PassThru
    # Write these as DECIMAL. In Windows PowerShell 5.1 the literal 0x80000000 is
    # parsed as a signed Int32 first and overflows to -2147483648, so
    # [uint32]0x80000000 throws - and keep-awake would silently never engage.
    $ES_CONTINUOUS = [uint32]2147483648
    $ES_SYSTEM_REQUIRED = [uint32]1
    # ES_DISPLAY_REQUIRED matters MORE than ES_SYSTEM_REQUIRED on a Modern Standby
    # machine: it enters connected standby when the DISPLAY turns off, regardless of
    # the standby timeout. Holding the display on is what actually keeps the network
    # alive. (Two runs died despite ES_SYSTEM_REQUIRED alone.)
    $ES_DISPLAY_REQUIRED = [uint32]2
    if ($keepAwake::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_SYSTEM_REQUIRED -bor $ES_DISPLAY_REQUIRED) -eq 0) {
        Write-Log "Could not request keep-awake; the machine may sleep mid-run." 'WARN'
        $keepAwake = $null
    } else {
        Write-Log "Keep-awake active - Windows will not sleep while this run is in progress."
    }
} catch {
    Write-Log "Keep-awake unavailable ($($_.Exception.Message)); the machine may sleep mid-run." 'WARN'
    $keepAwake = $null
}

# ------------------------------------------- suspend idle standby ------------
# Modern Standby ignores SetThreadExecutionState, so the only reliable way to keep
# a multi-hour run alive is to set the idle standby timeout to "never" while it
# runs, then put the original values back.
#
# The originals are saved to .power-restore.json BEFORE anything is changed, and
# only written if that file does not already exist - so a run that gets killed
# cannot overwrite the real values with 0. The next run restores from it first.
$powerMarker = Join-Path $PSScriptRoot '.power-restore.json'

function Get-TimeoutPair {
    param([string]$SubGroup, [string]$Setting)
    $out = powercfg /q SCHEME_CURRENT $SubGroup $Setting 2>$null
    $ac = $null; $dc = $null
    foreach ($l in $out) {
        if ($l -match 'Current AC Power Setting Index:\s*0x([0-9a-fA-F]+)') { $ac = [Convert]::ToInt32($Matches[1], 16) }
        if ($l -match 'Current DC Power Setting Index:\s*0x([0-9a-fA-F]+)') { $dc = [Convert]::ToInt32($Matches[1], 16) }
    }
    if ($null -eq $ac -or $null -eq $dc) { return $null }
    # powercfg reports SECONDS but /change takes MINUTES.
    return @{ acMin = [int]($ac / 60); dcMin = [int]($dc / 60) }
}

# Both matter. On Modern Standby the MONITOR timeout is the one that actually
# triggers connected standby, so it must be held off too.
function Get-StandbyTimeouts {
    $sleep   = Get-TimeoutPair 'SUB_SLEEP' 'STANDBYIDLE'
    $monitor = Get-TimeoutPair 'SUB_VIDEO' 'VIDEOIDLE'
    if (-not $sleep) { return $null }
    return @{
        acMin = $sleep.acMin; dcMin = $sleep.dcMin
        monAcMin = if ($monitor) { $monitor.acMin } else { -1 }
        monDcMin = if ($monitor) { $monitor.dcMin } else { -1 }
    }
}

function Restore-StandbyTimeouts {
    if (-not (Test-Path $powerMarker)) { return }
    try {
        $saved = Get-Content $powerMarker -Raw | ConvertFrom-Json
        powercfg /change standby-timeout-ac $saved.acMin 2>$null | Out-Null
        powercfg /change standby-timeout-dc $saved.dcMin 2>$null | Out-Null
        if ($saved.monAcMin -ge 0) { powercfg /change monitor-timeout-ac $saved.monAcMin 2>$null | Out-Null }
        if ($saved.monDcMin -ge 0) { powercfg /change monitor-timeout-dc $saved.monDcMin 2>$null | Out-Null }
        Write-Log "Restored power timeouts (sleep AC $($saved.acMin)/DC $($saved.dcMin) min, screen AC $($saved.monAcMin)/DC $($saved.monDcMin) min)."
    } catch {
        Write-Log "Could not restore sleep timeouts: $_" 'WARN'
    }
    Remove-Item $powerMarker -Force -ErrorAction SilentlyContinue
}

# A previous run may have been killed before restoring - put things back first.
Restore-StandbyTimeouts

$orig = Get-StandbyTimeouts
if ($orig) {
    if (-not (Test-Path $powerMarker)) {
        $orig | ConvertTo-Json | Set-Content -Path $powerMarker -Encoding utf8
    }
    powercfg /change standby-timeout-ac 0 2>$null | Out-Null
    powercfg /change standby-timeout-dc 0 2>$null | Out-Null
    # The screen timeout is the one that really triggers Modern Standby.
    powercfg /change monitor-timeout-ac 0 2>$null | Out-Null
    powercfg /change monitor-timeout-dc 0 2>$null | Out-Null
    Write-Log "Sleep AND screen blanking disabled for this run (was sleep AC $($orig.acMin)/DC $($orig.dcMin) min, screen AC $($orig.monAcMin)/DC $($orig.monDcMin) min); restored when it ends."
} else {
    Write-Log "Could not read the current sleep timeouts - the machine may sleep mid-run." 'WARN'
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
        $yearLabel = if ($Year -ge 2000) { "$Year" } else { '(current year)' }
        Write-Log "Reporting on month $Month $yearLabel (explicit)"
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
    # Release the sleep block - ES_CONTINUOUS on its own clears the earlier request.
    if ($keepAwake) {
        try { [void]$keepAwake::SetThreadExecutionState([uint32]2147483648) } catch { }
    }
    # Always hand the machine's normal sleep behaviour back.
    Restore-StandbyTimeouts
}
