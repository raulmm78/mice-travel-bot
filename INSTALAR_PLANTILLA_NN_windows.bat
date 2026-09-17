@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0instalar_plantilla_nn_windows.ps1" (
  echo Extrae el ZIP completo antes de ejecutarlo.
  pause
  exit /b 1
)
if not exist "%~dp0plantilla_nn_2026_09_17.xlsx" (
  echo Falta la plantilla NN en esta carpeta.
  pause
  exit /b 1
)

echo Cierra MICE Travel Bot antes de instalar la plantilla.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar_plantilla_nn_windows.ps1"
if errorlevel 1 (
  echo.
  echo No se completo la instalacion. Revisa el mensaje anterior.
  pause
  exit /b 1
)

echo.
echo Plantilla instalada y programa actualizado. Ya puedes abrir MICE Travel Bot.
pause
