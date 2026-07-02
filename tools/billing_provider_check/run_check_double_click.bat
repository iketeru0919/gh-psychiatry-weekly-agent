@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_check.ps1" -Interactive
echo.
echo Finished. Press any key to close.
pause >nul
