@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  set "PY=python"
)
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

set "PORT=8765"
if exist "config\.env" (
  for /f "usebackq tokens=1,* delims==" %%A in ("config\.env") do (
    if /i "%%A"=="EMAIL_AGENT_PORT" set "PORT=%%B"
  )
)

set "LOG_FILE=%TEMP%\mice_travel_bot_app.log"

powershell -NoProfile -Command "try { $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 1 http://127.0.0.1:%PORT%/; if ($r.Content -like '*MICE Travel Bot*') { exit 0 }; exit 1 } catch { exit 1 }" >nul 2>nul
if not errorlevel 1 (
  start "" "http://127.0.0.1:%PORT%/"
  exit /b 0
)

echo Buscando actualizaciones y arrancando MICE Travel Bot...
%PY% -u app\update_before_start.py > "%LOG_FILE%" 2>&1

start "MICE Travel Bot" /min "%~dp0app\run_panel_windows.bat"

for /l %%I in (1,1,30) do (
  powershell -NoProfile -Command "try { $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 1 http://127.0.0.1:%PORT%/; exit 0 } catch { exit 1 }" >nul 2>nul
  if not errorlevel 1 (
    start "" "http://127.0.0.1:%PORT%/"
    exit /b 0
  )
  timeout /t 1 /nobreak >nul
)

echo No se pudo abrir el panel.
echo Revisa el log: %LOG_FILE%
if not defined CI pause
exit /b 1
