$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
while ($root -and -not (Test-Path (Join-Path $root "MICE Travel Bot_windows_con_claves_20260913\MICE Travel Bot\app\process_emails.py"))) {
    $parent = Split-Path -Parent $root
    if (-not $parent -or $parent -eq $root) { break }
    $root = $parent
}
$install = Join-Path $root "MICE Travel Bot_windows_con_claves_20260913\MICE Travel Bot"
$log = Join-Path $root "INSTALACION_MICE_TRAVEL_BOT_v6.log"
$codeSource = Join-Path $PSScriptRoot "process_emails.py"
$launcherSource = Join-Path $PSScriptRoot "abrir_panel_windows.bat"
$codeTarget = Join-Path $install "app\process_emails.py"
$launcherTarget = Join-Path $install "abrir_panel_windows.bat"
$backupDir = Join-Path $install "_bot_backups_code"
$expectedCodeHash = "8A54C3DB49986F15D4621596427EC2EDBF36F4DFD16F5841950A7CB7FA343CF2"
$expectedLauncherHash = "F7E1750537EFDEA5E1BEB3DCA38EB133136FCD8B0C28E9EFAA9B26B2D20B890C"
$replaced = $false

function Get-PortableHash([string]$path) {
    $content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8).Replace("`r`n", "`n")
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-", "")
    } finally {
        $sha.Dispose()
    }
}

"Instalacion v6 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Set-Content -Path $log -Encoding UTF8
try {
    $listeningBotPorts = 8765..8784 | Where-Object {
        Get-NetTCPConnection -LocalPort $_ -State Listen -ErrorAction SilentlyContinue
    }
    if ($listeningBotPorts) {
        throw "El panel sigue abierto. Cierra MICE Travel Bot antes de instalar v6."
    }
    if (-not (Test-Path $codeTarget) -or -not (Test-Path $launcherTarget)) {
        throw "No encuentro la instalacion del bot en $install"
    }
    if (-not (Test-Path $codeSource) -or -not (Test-Path $launcherSource)) {
        throw "Faltan archivos de la instalacion. Espera a que OneDrive termine de sincronizar la carpeta."
    }
    if ((Get-PortableHash $codeSource) -ne $expectedCodeHash) {
        throw "El codigo v6 no coincide con el archivo esperado. No se ha instalado."
    }
    if ((Get-PortableHash $launcherSource) -ne $expectedLauncherHash) {
        throw "El lanzador v6 no coincide con el archivo esperado. No se ha instalado."
    }

    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $codeBackup = Join-Path $backupDir "process_emails.$stamp.py"
    $launcherBackup = Join-Path $backupDir "abrir_panel_windows.$stamp.bat"
    Copy-Item $codeTarget $codeBackup
    Copy-Item $launcherTarget $launcherBackup
    "Copias del codigo anterior: $backupDir" | Add-Content -Path $log

    Copy-Item $codeSource $codeTarget -Force
    $replaced = $true
    Copy-Item $launcherSource $launcherTarget -Force
    if ((Get-PortableHash $codeTarget) -ne $expectedCodeHash) {
        throw "El codigo copiado no supera la verificacion."
    }
    if ((Get-PortableHash $launcherTarget) -ne $expectedLauncherHash) {
        throw "El lanzador copiado no supera la verificacion."
    }

    $python = Join-Path $install ".venv\Scripts\python.exe"
    if (-not (Test-Path $python)) {
        throw "No encuentro Python en la instalacion."
    }
    & $python -m py_compile $codeTarget 2>&1 | Out-String | Add-Content -Path $log
    if ($LASTEXITCODE -ne 0) { throw "El codigo no compila en este Windows." }
    & $python -c "import sys; sys.path.insert(0, sys.argv[1]); import process_emails; print(process_emails.APP_VERSION)" (Join-Path $install "app") 2>&1 | Out-String | Add-Content -Path $log
    if ($LASTEXITCODE -ne 0) { throw "La importacion del bot ha fallado." }
    "OK: MICE TRAVEL BOT v6 instalado y comprobado." | Add-Content -Path $log
    Write-Host "OK: MICE TRAVEL BOT v6 instalado y comprobado."
} catch {
    "ERROR: $($_.Exception.Message)" | Add-Content -Path $log
    if ($replaced) {
        Copy-Item $codeBackup $codeTarget -Force
        Copy-Item $launcherBackup $launcherTarget -Force
        "Se ha restaurado el codigo anterior." | Add-Content -Path $log
    }
    Write-Host "ERROR: $($_.Exception.Message)"
    exit 1
}
