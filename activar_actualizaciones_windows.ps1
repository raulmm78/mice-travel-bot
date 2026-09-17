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
Write-Host "Cierra el panel si esta abierto y vuelve a abrir MICE Travel Bot."
Write-Host "En ese arranque descargara automaticamente la version mas reciente."
