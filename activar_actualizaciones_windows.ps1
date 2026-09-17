$ErrorActionPreference = "Stop"

$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "MICE Travel Bot.lnk"
if (-not (Test-Path $shortcutPath)) {
    throw "No encuentro el acceso directo 'MICE Travel Bot' en el escritorio. Instala primero el bot."
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$installDir = Split-Path -Parent $shortcut.TargetPath
$updater = Join-Path $installDir "app\update_before_start.py"
if (-not (Test-Path $updater)) {
    throw "No encuentro el actualizador en $installDir"
}

$download = Join-Path $env:TEMP "mice_update_before_start.py"
$source = "https://raw.githubusercontent.com/raulmm78/mice-travel-bot/main/app/update_before_start.py"
Invoke-WebRequest -UseBasicParsing -Uri $source -OutFile $download
if (-not (Select-String -Path $download -Pattern "def update_zip_installation" -Quiet)) {
    throw "La descarga no contiene el actualizador esperado. No se ha cambiado la instalacion."
}

$backup = "$updater.antes_github"
Copy-Item $updater $backup -Force
Copy-Item $download $updater -Force
Write-Host "Actualizaciones desde GitHub activadas en $installDir"

$python = Join-Path $installDir ".venv\Scripts\python.exe"
if (Test-Path $python) {
    & $python -u $updater
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 -u $updater
} else {
    & python -u $updater
}
$installedCommit = Get-Content (Join-Path $installDir "config\.github_commit") -ErrorAction SilentlyContinue | Select-Object -First 1
$latestCommit = (Invoke-RestMethod -Uri "https://api.github.com/repos/raulmm78/mice-travel-bot/commits/main" -Headers @{ "User-Agent" = "MICE-Travel-Bot-Updater" }).sha
if ($installedCommit -ne $latestCommit) {
    throw "No se pudo verificar la ultima version. Revisa la conexion a GitHub y vuelve a ejecutar este archivo."
}
Write-Host "Version actual descargada. Abre MICE Travel Bot desde el escritorio."
