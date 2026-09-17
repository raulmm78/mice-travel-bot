$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$install = Join-Path $root "MICE Travel Bot_windows_con_claves_20260913\MICE Travel Bot"
$target = Join-Path $install "app\process_emails.py"
$source = Join-Path $PSScriptRoot "process_emails.py"
$expectedHash = "8A54C3DB49986F15D4621596427EC2EDBF36F4DFD16F5841950A7CB7FA343CF2"

function Get-PortableHash([string]$path) {
    $content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8).Replace("`r`n", "`n")
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-", "") }
    finally { $sha.Dispose() }
}

try {
    if (-not (Test-Path $target) -or -not (Test-Path $source)) {
        throw "Faltan archivos del bot o de la version v6. Espera a que OneDrive sincronice."
    }
    if ((Get-PortableHash $source) -ne $expectedHash) {
        throw "La version v6 no se ha sincronizado correctamente."
    }

    $botProcessIds = @()
    foreach ($port in 8765..8784) {
        $connection = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
            Where-Object { $_.LocalAddress -in @("127.0.0.1", "0.0.0.0", "::1") } |
            Select-Object -First 1
        if (-not $connection) { continue }
        $process = Get-CimInstance Win32_Process -Filter "ProcessId = $($connection.OwningProcess)" -ErrorAction SilentlyContinue
        if (-not $process -or $process.CommandLine -notmatch "process_emails\.py") {
            throw "El puerto $port pertenece a otro programa. No se cerrara por seguridad."
        }
        try {
            $status = Invoke-RestMethod -Uri "http://127.0.0.1:$port/api/status" -TimeoutSec 3
        } catch {
            throw "No se puede comprobar si MICE Travel Bot esta guardando un Excel. Espera a que responda."
        }
        if ($status.phase -match "Guardando|Extrayendo|Importando|Comprobando") {
            throw "El bot anterior sigue procesando ($($status.phase)). Espera antes de abrir v6."
        }
        $botProcessIds += $connection.OwningProcess
    }
    foreach ($processId in ($botProcessIds | Select-Object -Unique)) {
        Write-Host "Cerrando MICE Travel Bot anterior..."
        Stop-Process -Id $processId -Force -ErrorAction Stop
    }
    if ($botProcessIds.Count -gt 0) {
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            $occupied = @(8765..8784 | ForEach-Object {
                Get-NetTCPConnection -LocalPort $_ -State Listen -ErrorAction SilentlyContinue
            })
            if ($occupied.Count -eq 0) { break }
            Start-Sleep -Milliseconds 500
        }
    }
    if ((Get-PortableHash $target) -ne $expectedHash) {
        Write-Host "Aplicando MICE Travel Bot v6..."
        & (Join-Path $PSScriptRoot "instalar_v6.ps1")
        if ($LASTEXITCODE -ne 0) { throw "No se ha completado la actualizacion v6." }
    }
    Write-Host "MICE Travel Bot v6 listo."
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    exit 1
}
