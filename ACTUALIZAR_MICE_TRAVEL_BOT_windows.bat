@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0activar_actualizaciones_windows.ps1" (
  echo Extrae el ZIP completo antes de ejecutar este archivo.
  pause
  exit /b 1
)

echo ================================================
echo Actualizacion MICE Travel Bot desde GitHub
echo ================================================
echo.
echo Cierra primero MICE Travel Bot si esta abierto.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0activar_actualizaciones_windows.ps1"
if errorlevel 1 (
  echo.
  echo No se completo la actualizacion. Revisa el mensaje anterior.
  pause
  exit /b 1
)

echo.
echo Actualizacion terminada. Ya puedes abrir MICE Travel Bot.
pause
