@echo off
setlocal
cd /d "%~dp0.."

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  set "PY=python"
)
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

set "LOG_FILE=%TEMP%\mice_travel_bot_app.log"
%PY% -u app\process_emails.py --dashboard >> "%LOG_FILE%" 2>&1
