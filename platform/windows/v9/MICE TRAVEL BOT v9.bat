@echo off
setlocal
set "UPDATE=%~dp0_bot_v9\preparar_v9.ps1"
set "BOT=%~dp0MICE Travel Bot_windows_con_claves_20260913\MICE Travel Bot\abrir_panel_windows.bat"
if not exist "%UPDATE%" (
  echo Faltan los archivos de MICE Travel Bot v9. Espera a OneDrive.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%UPDATE%"
if errorlevel 1 (
  echo No se abre el bot porque la actualizacion no se ha completado.
  pause
  exit /b 1
)
if not exist "%BOT%" (
  echo No encuentro el programa MICE Travel Bot.
  pause
  exit /b 1
)
call "%BOT%"
