@echo off
setlocal
cd /d "%~dp0"

echo ================================================
echo Instalador MICE Travel Bot
echo ================================================
echo.

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

where git >nul 2>nul
if not %errorlevel%==0 (
  echo.
  echo No se ha encontrado Git. Intentando instalar Git para las actualizaciones...
  where winget >nul 2>nul
  if %errorlevel%==0 (
    winget install --id Git.Git -e --accept-package-agreements --accept-source-agreements
  ) else (
    echo No se ha podido instalar Git automaticamente. El bot funcionara, pero sin autoactualizacion desde GitHub.
  )
)

echo Instalando dependencias...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt

if not exist ".env" (
  copy ".env.example" ".env" >nul
  echo.
  echo Se ha creado el archivo .env.
  echo Editalo con las claves de correo, OpenAI y rutas de OneDrive antes de arrancar.
) else (
  echo.
  echo Ya existe .env. No se ha tocado.
)

echo.
echo Creando accesos directos...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0crear_accesos_windows.ps1"

echo.
echo Instalacion terminada.
echo.
echo Siguiente paso:
echo 1. Edita el archivo .env con las claves y rutas de OneDrive.
echo 2. Abre "MICE Travel Bot" desde el escritorio.
if not defined CI pause
