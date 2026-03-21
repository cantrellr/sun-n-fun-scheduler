@echo off
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT_DIR%Start-SunNFunScheduler.ps1"
