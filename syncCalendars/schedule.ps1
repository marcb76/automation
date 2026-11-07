<#
.SYNOPSIS
     Script Name : schedule.ps1
     Version     : 1.0.4
     Created     : November 2026
     Author      : Marc Bonet
     Description : Task Scheduler installer script for syncCalendars automation utility.
     Objective   : Register or update the hourly task in Windows Task Scheduler.
#>

Clear-Host

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Script Name : schedule.ps1" -ForegroundColor Cyan
Write-Host " Version     : 1.0.4" -ForegroundColor Cyan
Write-Host " Created     : November 2026" -ForegroundColor Cyan
Write-Host " Author      : Marc Bonet" -ForegroundColor Cyan
Write-Host " Description : Task Scheduler installer script for" -ForegroundColor Cyan
Write-Host "               syncCalendars automation utility." -ForegroundColor Cyan
Write-Host " Objective   : Register or update the hourly task in Windows" -ForegroundColor Cyan
Write-Host "               Task Scheduler with description." -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

$TaskName   = "mb_syncCalendars"
$TaskDesc   = "Hourly bidirectional synchronization between Outlook Desktop Calendar and Google Calendar"
$ScriptPath = "C:\Users\marcb\OneDrive\Personal Managed\DevEnv\automation\syncCalendars\automate.ps1"

Write-Host "Checking if task '$TaskName' already exists..." -ForegroundColor Yellow
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Task '$TaskName' already exists. Removing existing task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

Write-Host "Registering task '$TaskName' in Windows Task Scheduler..." -ForegroundColor Cyan
Write-Host "Script Target: $ScriptPath"
Write-Host "Description  : $TaskDesc"
Write-Host "Schedule     : Hourly starting at 08:00 AM (10-year indefinite repetition)"
Write-Host "------------------------------------------------------------------------------" -ForegroundColor Cyan

$Trigger = New-ScheduledTaskTrigger -Once -At "08:00 AM" -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Days 3650)
$Action  = New-ScheduledTaskAction -Execute "powershell.exe" `
                                  -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

try {
    # Using 'Limited' which is the correct PowerShell enum for standard user privileges
    Register-ScheduledTask -TaskName $TaskName `
                           -Action $Action `
                           -Trigger $Trigger `
                           -Description $TaskDesc `
                           -RunLevel Limited `
                           -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable) `
                           -ErrorAction Stop

    Write-Host ""
    Write-Host "==============================================================================" -ForegroundColor Green
    Write-Host " Task successfully registered with standard user privileges!" -ForegroundColor Green
    Write-Host "==============================================================================" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "ERROR: Failed to register the task: $_" -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"