@echo off
setlocal
set "ROOT=%~dp0"
set "INSTALL=%ROOT%MICE Travel Bot_windows_con_claves_20260913\MICE Travel Bot"
set "PYTHON=%INSTALL%\.venv\Scripts\python.exe"
set "ENV=%INSTALL%\config\.env"
set "CHECK=%ROOT%_bot_v9\probar_correo_produccion.py"
if not exist "%PYTHON%" (
  echo ERROR: No se encuentra Python en la instalacion del bot.
  pause
  exit /b 1
)
if not exist "%CHECK%" (
  echo ERROR: Falta el comprobador. Espera a que OneDrive sincronice.
  pause
  exit /b 1
)
"%PYTHON%" "%CHECK%" "%ENV%"
set "RESULT=%errorlevel%"
pause
exit /b %RESULT%
