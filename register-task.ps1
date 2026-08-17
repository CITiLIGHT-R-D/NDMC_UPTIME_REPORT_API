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
    [switch]$Unregister,
    [int]$DayOfMonth = 2,
    [string]$AtTime = '01:00'
)

$ErrorActionPreference = 'Stop'
$TaskName = 'NDMC Monthly Reports'
$scriptDir = $PSScriptRoot
$wrapper = Join-Path $scriptDir 'run-monthly.ps1'

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
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed scheduled task '$TaskName'." -ForegroundColor Green
    } else {
        Write-Host "No scheduled task named '$TaskName' found."
    }
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
$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$wrapper`"" `
    -WorkingDirectory $scriptDir

$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth $DayOfMonth -At $AtTime

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -WakeToRun `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 6) `
    -MultipleInstances IgnoreNew `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 30)

# S4U runs the task whether or not the user is logged on, WITHOUT storing a
# password. Outbound HTTPS (the portal and Gmail) works fine under it.
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType S4U `
    -RunLevel Limited

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Task already exists — replacing it."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Generates the NDMC uptime and operational hour reports for the month that just ended, verifies every zone is complete, and emails them to the recipients in recipients.txt." | Out-Null

Write-Host ""
Write-Host "Registered '$TaskName'." -ForegroundColor Green
Write-Host "  Runs   : day $DayOfMonth of each month at $AtTime"
Write-Host "  Reports: the month that just ended (2 Aug -> July, 2 Jan -> December)"
Write-Host "  As     : $env:USERDOMAIN\$env:USERNAME (whether logged on or not, no password stored)"
Write-Host "  Script : $wrapper"
Write-Host "  Logs   : $(Join-Path $scriptDir 'Logs')"
Write-Host ""
Write-Host "Missed runs: if this PC is off at $AtTime, Windows runs the job as soon" -ForegroundColor Cyan
Write-Host "as the machine is next available, so the month is not skipped." -ForegroundColor Cyan
Write-Host ""
Write-Host "Test it now without waiting:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "  Get-ScheduledTaskInfo -TaskName '$TaskName'"
$next = (Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo).NextRunTime
if ($next) { Write-Host ""; Write-Host "Next run: $next" -ForegroundColor Green }
