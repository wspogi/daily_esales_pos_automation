@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Running Daily E-Sales Report POS Automation...
echo ============================================================

REM Manual month example:
REM py main.py 2026-04

py main.py

if errorlevel 1 (
    echo.
    echo ERROR: Report generation failed. Check logs folder.
    pause
    exit /b 1
)

echo.
echo DONE: Report generated successfully.
pause
