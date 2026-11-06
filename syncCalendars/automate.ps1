<#
.SYNOPSIS
     Script Name : automate.ps1
     Version     : 1.0.0
     Created     : November 2026
     Author      : Marc Bonet
     Description : Wrapper execution script for syncCalendars utility.
     Objective   : Activates the local venv and runs the Python script with --now.
#>

# Move to the script's directory to ensure correct relative paths
Set-Location -Path $PSScriptRoot

Write-Host "Starting calendar synchronization..." -ForegroundColor Cyan

# Execute the Python script using the local virtual environment interpreter
& ".\venv\Scripts\python.exe" "syncCalendars.py" --now

if ($LASTEXITCODE -eq 0) {
    Write-Host "Synchronization completed successfully." -ForegroundColor Green
} else {
    Write-Host "Synchronization finished with errors (Exit Code: $LASTEXITCODE)." -ForegroundColor Red
}
