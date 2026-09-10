$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
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
$envFile = Join-Path $baseDir ".env"

New-Shortcut -Name "MICE Travel Bot" -Target $openBot
New-Shortcut -Name "Cerrar MICE Travel Bot" -Target $closeBot
New-Shortcut -Name "Configurar MICE Travel Bot" -Target "notepad.exe" -Arguments "`"$envFile`""

Write-Host "Accesos directos creados en el escritorio."
