@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  set "PY=python"
)

%PY% process_emails.py --diagnostico
echo.
if errorlevel 1 (
  echo Hay puntos que revisar antes de ejecutar el bot.
) else (
  echo Todo listo para probar.
)
if not defined CI pause
