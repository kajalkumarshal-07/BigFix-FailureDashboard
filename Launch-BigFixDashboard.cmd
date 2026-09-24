@echo off
rem BigFix Failed Device Dashboard launcher
rem Starts PowerShell hidden (GUI only); cmd exits immediately - no leftover console.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0BigFix-FailureDashboard.ps1" %*
exit /b 0
