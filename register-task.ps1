<#
================================================================================
  REGISTER THE MONTHLY SCHEDULED TASK
================================================================================
  Creates a Windows Task Scheduler job that runs run-monthly.ps1 on the 2nd of
  every month at 01:00, generating and emailing the reports for the month that
  just ended.

  RUN THIS ONCE, FROM AN ADMIN POWERSHELL:
      cd "D:\NDMC UPTIME REPORT API"
      .\register-task.ps1

  To remove it:
      .\register-task.ps1 -Unregister

  To check it, or run it right now without waiting for the 2nd:
      Get-ScheduledTask  -TaskName 'NDMC Monthly Reports' | Format-List
      Get-ScheduledTaskInfo -TaskName 'NDMC Monthly Reports'
      Start-ScheduledTask   -TaskName 'NDMC Monthly Reports'

  WHY THE SETTINGS BELOW MATTER ON A PC (rather than an always-on server):
    - StartWhenAvailable  : a desktop is often off at 01:00. Without this the
                            run is simply skipped and the month is missed. With
                            it, Windows runs the job at the next opportunity.
    - WakeToRun           : wakes the machine from sleep to run.
    - ExecutionTimeLimit  : the full run takes ~2h; the default limit would
                            kill it partway.
    - Battery settings    : on a laptop, tasks are stopped/blocked on battery
                            by default, which would abort a run mid-way.
================================================================================
#>

[CmdletBinding()]
param(
    # Direct   : Task Scheduler triggers the run itself on day $DayOfMonth. One
    #            OS-level entry, nothing running in between. Simplest.
    # Scheduler: Task Scheduler only starts scheduler.js at boot; node-cron owns
    #            the monthly trigger and re-checks on startup whether the month's
    #            reports actually exist, retrying if a previous run failed.
    [ValidateSet('Direct', 'Scheduler')]
    [string]$Mode = 'Scheduler',
    [switch]$Unregister,
    [int]$DayOfMonth = 2,
    [string]$AtTime = '01:00'
)

$ErrorActionPreference = 'Stop'
$TaskName = if ($Mode -eq 'Scheduler') { 'NDMC Report Scheduler' } else { 'NDMC Monthly Reports' }
$scriptDir = $PSScriptRoot
$wrapper = Join-Path $scriptDir 'run-monthly.ps1'
$scheduler = Join-Path $scriptDir 'scheduler.js'

# --- admin check ---
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This must be run from an ADMIN PowerShell window." -ForegroundColor Red
    Write-Host "  Right-click PowerShell -> 'Run as administrator', then re-run this script."
    exit 1
}

# --- remove ---
if ($Unregister) {
    # Remove BOTH task names, so switching modes can never leave the old one
    # firing alongside the new one and running two jobs against the portal.
    $removed = 0
    foreach ($n in @('NDMC Monthly Reports', 'NDMC Report Scheduler')) {
        if (Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $n -Confirm:$false
            Write-Host "Removed scheduled task '$n'." -ForegroundColor Green
            $removed++
        }
    }
    if (-not $removed) { Write-Host "No NDMC scheduled tasks found." }
    exit 0
}

# --- preflight ---
if (-not (Test-Path $wrapper)) { Write-Host "run-monthly.ps1 not found next to this script." -ForegroundColor Red; exit 1 }
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { Write-Host "Node.js is not on PATH." -ForegroundColor Red; exit 1 }

$cfg = Join-Path $scriptDir 'config.env'
if (-not (Test-Path $cfg)) {
    Write-Host "config.env not found — the task would fail on its first run." -ForegroundColor Yellow
    Write-Host "  Copy config.env.example to config.env and fill in the portal + mail settings first."
    Write-Host ""
    $reply = Read-Host "Register the task anyway? (y/N)"
    if ($reply -notmatch '^[Yy]') { exit 1 }
}

# --- build the task ---
if ($Mode -eq 'Scheduler') {
    if (-not (Test-Path $scheduler)) { Write-Host "scheduler.js not found." -ForegroundColor Red; exit 1 }
    $nodeExe = (Get-Command node).Source
    # Boot-start only: node-cron inside scheduler.js owns the monthly timing.
    $action  = New-ScheduledTaskAction -Execute $nodeExe -Argument "`"$scheduler`"" -WorkingDirectory $scriptDir
    $trigger = New-ScheduledTaskTrigger -AtStartup
    # A long-running process must not be killed part-way, and must come back if it dies.
    $extraSettings = @{ ExecutionTimeLimit = ([TimeSpan]::Zero); RestartCount = 3; RestartInterval = (New-TimeSpan -Minutes 5) }
} else {
    $action  = New-ScheduledTaskAction `
        -Execute 'powershell.exe' `
        -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$wrapper`"" `
        -WorkingDirectory $scriptDir
    $trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth $DayOfMonth -At $AtTime
    $extraSettings = @{ ExecutionTimeLimit = (New-TimeSpan -Hours 6); RestartCount = 2; RestartInterval = (New-TimeSpan -Minutes 30) }
}

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -WakeToRun `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit $extraSettings.ExecutionTimeLimit `
    -RestartCount $extraSettings.RestartCount `
    -RestartInterval $extraSettings.RestartInterval

# S4U runs the task whether or not the user is logged on, WITHOUT storing a
# password. Outbound HTTPS (the portal and Gmail) works fine under it.
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType S4U `
    -RunLevel Limited

# Clear BOTH names first: having Direct and Scheduler registered together would
# run the job twice against a portal that already struggles with one run.
foreach ($n in @('NDMC Monthly Reports', 'NDMC Report Scheduler')) {
    if (Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue) {
        Write-Host "Removing existing task '$n'."
        Unregister-ScheduledTask -TaskName $n -Confirm:$false
    }
}

$desc = if ($Mode -eq 'Scheduler') {
    "Starts scheduler.js at boot. node-cron fires the monthly NDMC report run, and on every startup it re-checks whether the month's reports actually exist, retrying if a previous run failed."
} else {
    "Generates the NDMC uptime and operational hour reports for the month that just ended, verifies every zone is complete, and emails them to the recipients in recipients.txt."
}

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description $desc | Out-Null

Write-Host ""
Write-Host "Registered '$TaskName'  (mode: $Mode)" -ForegroundColor Green
if ($Mode -eq 'Scheduler') {
    Write-Host "  Trigger: at system startup — scheduler.js then stays running"
    Write-Host "  Timing : node-cron, SCHEDULE_CRON (default '0 1 2 * *' = 01:00 on the 2nd)"
    Write-Host "  Catchup: on every startup it checks whether the month's reports exist,"
    Write-Host "           and runs them if they are missing or a previous attempt failed"
    Write-Host "  Runs   : $((Get-Command node).Source) `"$scheduler`""
} else {
    Write-Host "  Trigger: day $DayOfMonth of each month at $AtTime"
    Write-Host "  Catchup: Windows StartWhenAvailable (runs when the PC is next on)"
    Write-Host "  Runs   : $wrapper"
}
Write-Host "  Period : the month that just ended (2 Aug -> July, 2 Jan -> December)"
Write-Host "  As     : $env:USERDOMAIN\$env:USERNAME (whether logged on or not, no password stored)"
Write-Host "  Logs   : $(Join-Path $scriptDir 'Logs')"
Write-Host ""
Write-Host "Start it now without rebooting:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
if ($Mode -eq 'Scheduler') {
    Write-Host ""
    Write-Host "Check what the scheduler thinks:" -ForegroundColor Yellow
    Write-Host "  node scheduler.js --status"
    Write-Host "  node scheduler.js --run-now      # force a run immediately"
}
$next = (Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo).NextRunTime
if ($next) { Write-Host ""; Write-Host "Next run: $next" -ForegroundColor Green }
