$ErrorActionPreference = "Stop"

$shortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "MICE Travel Bot.lnk"
if (-not (Test-Path $shortcutPath)) {
    throw "No encuentro MICE Travel Bot en el escritorio de Windows."
}
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$installDir = Split-Path -Parent $shortcut.TargetPath
$source = Join-Path $PSScriptRoot "plantilla_nn_2026_09_17.xlsx"
$destination = Join-Path $installDir "app\assets\plantilla_nn_2026_09_17.xlsx"
if (-not (Test-Path $source)) {
    throw "Falta la plantilla dentro del ZIP. Extrae el ZIP completo."
}
if (-not (Test-Path (Join-Path $installDir "app\process_emails.py"))) {
    throw "No encuentro la instalacion del bot en $installDir"
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
if (Test-Path $destination) {
    $sourceHash = (Get-FileHash $source -Algorithm SHA256).Hash
    $destinationHash = (Get-FileHash $destination -Algorithm SHA256).Hash
    if ($sourceHash -ne $destinationHash) {
        $backup = "$destination.antes_$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Copy-Item $destination $backup
        Copy-Item $source $destination -Force
    }
} else {
    Copy-Item $source $destination
}

Write-Host "Plantilla NN instalada en $destination"
& (Join-Path $PSScriptRoot "activar_actualizaciones_windows.ps1")
