@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Testing SQL Server connection...
echo ============================================================

py test_connection.py

pause
