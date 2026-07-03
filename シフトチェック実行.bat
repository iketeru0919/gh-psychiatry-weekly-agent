@echo off
rem ============================================
rem  Shift Weekly Checker - double click to run
rem ============================================
chcp 65001 > nul
cd /d "%~dp0"

where py > nul 2>&1
if %errorlevel% == 0 (
    py -3 -m shift_checker
) else (
    python -m shift_checker
)

echo.
pause
