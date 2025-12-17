@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -NoProfile -NoExit -File "Start-M365Manager.ps1"
