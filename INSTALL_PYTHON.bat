@echo off
setlocal

echo ============================================================
echo Python Installer Helper
echo ============================================================
echo This will open the official Python download page.
echo Download Python 3.x, run installer, and CHECK:
echo     Add python.exe to PATH
echo ============================================================

start https://www.python.org/downloads/windows/

echo.
echo After installing Python, reopen CMD and run:
echo     py --version
echo     INSTALL_REQUIREMENTS.bat
pause
