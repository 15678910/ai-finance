@echo off
chcp 949 >nul 2>&1

set "PYTHONIOENCODING=utf-8"
set "WORK_DIR=%~dp0"

cd /d "%WORK_DIR%"

if not exist "output\logs" mkdir "output\logs"

set "LOGDATE=%date:~0,4%%date:~5,2%%date:~8,2%"

python "%WORK_DIR%\daily_sector_analysis.py" >> "output\logs\daily_run_%LOGDATE%.log" 2>&1
