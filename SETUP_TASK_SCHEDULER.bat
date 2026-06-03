@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Setup Windows Task Scheduler
echo Daily E-Sales Report POS Automation
echo ============================================================
echo.
echo This creates a MONTHLY task every 1st day at 08:00 AM.
echo The script defaults to previous month report.
echo.

set TASK_NAME=Daily E-Sales Report POS Automation
set BAT_PATH=%~dp0RUN_DAILY_E_SALES_REPORT.bat

schtasks /Create ^
 /TN "%TASK_NAME%" ^
 /TR "\"%BAT_PATH%\"" ^
 /SC MONTHLY ^
 /D 1 ^
 /ST 08:00 ^
 /F

if errorlevel 1 (
    echo.
    echo ERROR: Failed to create scheduled task.
    echo Try running this BAT as Administrator.
    pause
    exit /b 1
)

echo.
echo DONE: Scheduled task created successfully.
echo Task Name: %TASK_NAME%
pause
