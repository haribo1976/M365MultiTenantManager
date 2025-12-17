@echo off
cd /d "%~dp0"
pwsh.exe -ExecutionPolicy Bypass -NoProfile -File "Start-M365Manager.ps1"
pause
