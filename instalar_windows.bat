@echo off
setlocal
cd /d "%~dp0"

echo ================================================
echo Instalador MICE Travel Bot
echo ================================================
echo.

echo Comprobando carpeta de instalacion...
echo %CD% | findstr /i "\\\\.zip\\" >nul
if %errorlevel%==0 (
  echo.
  echo Parece que has abierto el instalador desde dentro del ZIP.
  echo.
  echo Solucion:
  echo 1. Cierra esta ventana.
  echo 2. Boton derecho sobre el ZIP.
  echo 3. Pulsa "Extraer todo...".
  echo 4. Entra en la carpeta extraida "MICE Travel Bot".
  echo 5. Ejecuta de nuevo instalar_windows.bat.
  echo.
  explorer "%~dp0"
  if not defined CI pause
  exit /b 1
)

if not exist "app\requirements.txt" (
  echo.
  echo No encuentro app\requirements.txt.
  echo Esto suele pasar si el ZIP no se ha extraido correctamente o si estas dentro de una carpeta incompleta.
  echo.
  echo Comprueba que en esta carpeta existan:
  echo - app
  echo - config
  echo - docs
  echo - instalar_windows.bat
  echo.
  echo Si estas viendo el contenido desde el ZIP, usa "Extraer todo..." primero.
  echo.
  explorer "%CD%"
  if not defined CI pause
  exit /b 1
)

if not exist "app\process_emails.py" (
  echo.
  echo No encuentro app\process_emails.py.
  echo La carpeta del bot esta incompleta. Extrae de nuevo el ZIP completo.
  echo.
  explorer "%CD%"
  if not defined CI pause
  exit /b 1
)

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  set "PY=python"
)

%PY% --version >nul 2>nul
if not %errorlevel%==0 (
  echo No se ha encontrado Python. Intentando instalar Python...
  where winget >nul 2>nul
  if not %errorlevel%==0 (
    echo.
    echo No se ha encontrado winget en este PC.
    echo Instala Python 3 manualmente desde:
    echo https://www.python.org/downloads/windows/
    echo.
    echo Importante: durante la instalacion marca "Add python.exe to PATH".
    if not defined CI pause
    exit /b 1
  )
  winget install --id Python.Python.3.12 -e --accept-package-agreements --accept-source-agreements
  where py >nul 2>nul
  if %errorlevel%==0 (
    set "PY=py -3"
  ) else (
    set "PY=python"
  )
  %PY% --version >nul 2>nul
  if not %errorlevel%==0 (
    echo.
    echo Python se ha instalado, pero Windows aun no lo encuentra en esta ventana.
    echo Cierra esta ventana y vuelve a ejecutar instalar_windows.bat.
    if not defined CI pause
    exit /b 1
  )
)

echo Instalando dependencias...
%PY% -m venv .venv
set "PY=.venv\Scripts\python.exe"
%PY% -m pip install --upgrade pip
%PY% -m pip install -r app\requirements.txt

if not exist "config" mkdir config
if not exist "config\.env" (
  copy "config\.env.example" "config\.env" >nul
  set "NEEDS_CONFIG=1"
  echo.
  echo Se ha creado el archivo config\.env.
  echo Editalo con las claves de correo, OpenAI y rutas de OneDrive antes de arrancar.
) else (
  echo.
  echo Ya existe config\.env. Se conservaron sus claves y rutas.
)

echo.
echo Creando accesos directos...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0platform\windows\crear_accesos_windows.ps1"

echo.
echo Instalacion terminada.
echo.
echo Siguiente paso:
if defined NEEDS_CONFIG (
  echo 1. Edita config\.env con las claves y rutas de OneDrive.
  echo 2. Ejecuta diagnostico_windows.bat.
  echo 3. Abre "MICE Travel Bot" desde el escritorio.
) else (
  echo 1. Ejecuta diagnostico_windows.bat.
  echo 2. Abre "MICE Travel Bot" desde el escritorio.
)
if not defined CI pause
