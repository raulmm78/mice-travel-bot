$ErrorActionPreference = "Stop"

$repoUrl = "https://github.com/raulmm78/mice-travel-bot.git"
$installDir = Join-Path $env:LOCALAPPDATA "MICETravelBot"

function Command-Exists {
    param([string]$Command)
    $null -ne (Get-Command $Command -ErrorAction SilentlyContinue)
}

function Install-WithWinget {
    param(
        [string]$PackageId,
        [string]$Name
    )

    if (-not (Command-Exists "winget")) {
        throw "No se ha encontrado winget. Instala $Name manualmente y vuelve a ejecutar el instalador."
    }

    Write-Host "Instalando $Name..."
    winget install --id $PackageId -e --accept-package-agreements --accept-source-agreements
}

Write-Host "==============================================="
Write-Host "Instalador MICE Travel Bot desde GitHub"
Write-Host "==============================================="
Write-Host ""

if (-not (Command-Exists "py") -and -not (Command-Exists "python")) {
    Install-WithWinget -PackageId "Python.Python.3.12" -Name "Python 3"
}

if (-not (Command-Exists "git")) {
    Install-WithWinget -PackageId "Git.Git" -Name "Git"
}

if (-not (Command-Exists "git")) {
    throw "Git se ha instalado, pero Windows aun no lo encuentra. Cierra PowerShell y vuelve a ejecutar este instalador."
}

if (Test-Path (Join-Path $installDir ".git")) {
    Write-Host "Actualizando carpeta existente..."
    git -C $installDir pull --ff-only
} else {
    if (Test-Path $installDir) {
        $backup = "$installDir.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Write-Host "Ya existia una carpeta sin Git. Se guarda copia en $backup"
        Move-Item $installDir $backup
    }
    Write-Host "Descargando proyecto..."
    git clone $repoUrl $installDir
}

Write-Host ""
Write-Host "Ejecutando instalador local..."
& (Join-Path $installDir "instalar_windows.bat")

Write-Host ""
Write-Host "Listo. Revisa el archivo .env antes de arrancar el bot."
