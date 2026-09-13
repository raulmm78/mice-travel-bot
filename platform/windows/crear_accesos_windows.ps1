$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseDir = Split-Path -Parent (Split-Path -Parent $scriptDir)
$desktop = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell

function New-Shortcut {
    param(
        [string]$Name,
        [string]$Target,
        [string]$Arguments = "",
        [string]$Icon = ""
    )

    $shortcutPath = Join-Path $desktop "$Name.lnk"
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $Target
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = $baseDir
    if ($Icon) {
        $shortcut.IconLocation = $Icon
    }
    $shortcut.Save()
}

$openBot = Join-Path $baseDir "abrir_panel_windows.bat"
$closeBot = Join-Path $baseDir "cerrar_panel_windows.bat"
$diagnostic = Join-Path $baseDir "diagnostico_windows.bat"
$envFile = Join-Path $baseDir "config\.env"

New-Shortcut -Name "MICE Travel Bot" -Target $openBot
New-Shortcut -Name "Cerrar MICE Travel Bot" -Target $closeBot
New-Shortcut -Name "Diagnostico MICE Travel Bot" -Target $diagnostic
New-Shortcut -Name "Configurar MICE Travel Bot" -Target "notepad.exe" -Arguments "`"$envFile`""

Write-Host "Accesos directos creados en el escritorio."
