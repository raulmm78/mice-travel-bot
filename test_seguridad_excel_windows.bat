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

%PY% app\safety_test.py
echo.
if errorlevel 1 (
  echo ERROR - Revisa el mensaje anterior antes de usar el bot con Excels reales.
) else (
  echo OK - Puedes usar el bot sin borrar filas existentes.
)
if not defined CI pause
