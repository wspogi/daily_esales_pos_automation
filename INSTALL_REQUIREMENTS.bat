@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Installing Python requirements...
echo ============================================================

py -m pip install --upgrade pip
py -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Failed to install requirements.
    pause
    exit /b 1
)

echo.
echo DONE: Requirements installed successfully.
pause
