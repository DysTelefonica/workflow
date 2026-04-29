[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath
)

Set-StrictMode -Version Latest

function Get-ConfigField {
    param($obj, $field, $default = $null)
    if ($obj.PSObject.Properties[$field]) { return $obj.$field }
    return $default
}

# === Validar password ===
if (-not $env:ACCESS_SANDBOX_PW) {
    Write-Host "ERROR: La variable de entorno ACCESS_SANDBOX_PW no esta definida." -ForegroundColor Red
    Write-Host ""
    Write-Host "Para解决这个问题, ejecuta esto EN UNA TERMINAL NUEVA (PowerShell) y luego vuelve a ejecutar el script:" -ForegroundColor Yellow
    Write-Host "  setx ACCESS_SANDBOX_PW dpddpd" -ForegroundColor Cyan
    Write-Host "  (cerrar y abrir la terminal para que tome efecto)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "O ejecutalo desde el .bat (que lo setea automaticamente):" -ForegroundColor Yellow
    Write-Host "  .\sync.bat" -ForegroundColor Cyan
    exit 1
}
$backendPassword = $env:ACCESS_SANDBOX_PW

# === Cargar config ===
if (-not (Test-Path $ConfigPath)) {
    Write-Host "ERROR: No se encontro el archivo de configuracion: $ConfigPath" -ForegroundColor Red
    exit 1
}
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

$srcFolder = Get-ConfigField $config "sourceFolder"
$localSandboxPath = Get-ConfigField $config "localSandboxPath" "C:\00repos\datos"

if ($srcFolder -and -not (Test-Path $srcFolder)) {
    Write-Host "ERROR: La carpeta de origen no existe: $srcFolder" -ForegroundColor Red
    exit 1
}

function Write-Status {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    $old = $Host.UI.RawUI.ForegroundColor
    try {
        $Host.UI.RawUI.ForegroundColor = $Color
        Write-Host $Message
    } finally {
        $Host.UI.RawUI.ForegroundColor = $old
    }
}

function New-DaoDbEngine {
    $engineCandidates = @("DAO.DBEngine.160", "DAO.DBEngine.150", "DAO.DBEngine.140", "DAO.DBEngine.120", "DAO.DBEngine.36")
    foreach ($eng in $engineCandidates) {
        try {
            $ada = New-Object -ComObject $eng
            return $ada
        } catch { }
    }
    throw "No se pudo crear ningun DAO.DBEngine. Instala Microsoft Access o el Access Database Engine."
}

Write-Status "=== sync-backends.ps1 ===" -Color Green
Write-Status "Inicio: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Status "Sandbox local: $localSandboxPath"

# Crear sandbox si no existe
if (-not (Test-Path $localSandboxPath)) {
    New-Item -ItemType Directory -Path $localSandboxPath -Force | Out-Null
    Write-Status "Carpeta de sandbox creada." -Color Cyan
}

# Test conectividad a origen
function Test-PathAccessible {
    param([string]$FolderPath)
    try {
        Get-ChildItem $FolderPath -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# === Comprobar conectividad de red antes de tocar nada ===
$productionPaths = Get-ConfigField $config "productionPaths"
if ($productionPaths) {
    # Extraer el server UNC del primer path disponible
    $firstPath = $productionPaths[0]
    $uncServer = $firstPath -replace '^\\\\([^\\]+)\\.*', '$1'
    Write-Status "Verificando conectividad a \\$uncServer..." -Color Cyan
    if (-not (Test-Connection -ComputerName $uncServer -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        Write-Status "ERROR: No se puede acceder a \\$uncServer. Conectate a la VPN." -Color Red
        exit 1
    }
    Write-Status "Conectividad OK" -Color Green
} elseif ($srcFolder) {
    $srcAccessible = Test-PathAccessible -FolderPath $srcFolder
    if (-not $srcAccessible) {
        Write-Status "ERROR: No se puede acceder a la carpeta de origen: $srcFolder" -Color Red
        Write-Status "       Verifica la conexion de red." -Color Red
        exit 1
    }
}

$stamp = Get-Date -Format 'yyyyMMdd'

# === SEGURIDAD: Zip de lo que haya (no otros zips) ===
$itemsToZip = Get-ChildItem $localSandboxPath -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne '.zip' }
if ($itemsToZip -and @($itemsToZip).Count -gt 0) {
    $zipPath = Join-Path $localSandboxPath "backends_$stamp.zip"
    Write-Status "Detectados $(@($itemsToZip).Count) archivos en sandbox. Creando zip de seguridad..." -Color Yellow
    Compress-Archive -Path "$localSandboxPath\*" -DestinationPath $zipPath -Force
    Write-Status "Zip: $zipPath" -Color Gray
}

# === LIMPIEZA: Borrar todo menos zips ===
Write-Status "Limpiando sandbox (conservando zips)..." -Color Cyan
Get-ChildItem $localSandboxPath -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne '.zip' } | ForEach-Object {
    if ($_.PSIsContainer) {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
    } else {
        Remove-Item -LiteralPath $_.FullName -Force
    }
}

# === PHASE 1: Copiar accdb del origen ===
Write-Status "Fase 1: Copiando backends al sandbox..." -Color Cyan
$copiedFiles = New-Object System.Collections.ArrayList

if ($productionPaths -and @($productionPaths).Count -gt 0) {
    # Modo produccion: copiar archivos individuales por ruta UNC
    foreach ($srcPath in $productionPaths) {
        if (-not (Test-Path $srcPath)) {
            Write-Status "  SKIP (no existe): $srcPath" -Color Yellow
            continue
        }
        $fileName = [System.IO.Path]::GetFileName($srcPath)
        $dst = Join-Path $localSandboxPath $fileName
        Copy-Item $srcPath -Destination $dst -Force
        [void]$copiedFiles.Add($dst)
        Write-Status "  Copiado: $fileName" -Color Gray
    }
} elseif ($srcFolder) {
    # Modo dev: copiar todos los .accdb de una carpeta
    $srcAccdbs = Get-ChildItem $srcFolder -Filter '*.accdb' -ErrorAction SilentlyContinue
    if (-not $srcAccdbs) {
        Write-Status "WARNING: No se encontraron archivos .accdb en el origen." -Color Yellow
    }
    foreach ($f in $srcAccdbs) {
        $dst = Join-Path $localSandboxPath $f.Name
        Copy-Item $f.FullName -Destination $dst -Force
        [void]$copiedFiles.Add($dst)
        Write-Status "  Copiado: $($f.Name)" -Color Gray
    }
}
Write-Status "Fase 1 completada: $(@($copiedFiles).Count) archivos copiados." -Color Green

# === PHASE 2: Reescribir TableDef.Connect (two-pass: scan then apply) ===
Write-Status "Fase 2: Revinculando tablas vinculadas en $(@($copiedFiles).Count) backends..." -Color Cyan
$ada = New-DaoDbEngine

# Pass 1: detectar y planear
$backendPlan = @{}
foreach ($backendPath in $copiedFiles) {
    $fileName = [System.IO.Path]::GetFileName($backendPath)
    $db = $ada.OpenDatabase($backendPath, $false, $false, ";pwd=$backendPassword")
    $plan = @()

    foreach ($td in $db.TableDefs) {
        if ($td.Name -like 'MSys*') { continue }
        $connect = $td.Connect
        if ([string]::IsNullOrEmpty($connect)) { continue }
        # Buscar connects que apunten a otro .accdb (\\server\... o c:\...)
        if ($connect -match 'DATABASE=.*\\([^\\]+\.accdb)') {
            $linkedBackend = $matches[1]
            $destPath = Join-Path $localSandboxPath $linkedBackend
            $plan += [PSCustomObject]@{
                TableName     = $td.Name
                LinkedBackend = $linkedBackend
                DestPath      = $destPath
                BackendExists = (Test-Path $destPath)
            }
        }
    }
    $db.Close()
    $backendPlan[$fileName] = $plan
}

# Pass 2: aplicar (Connect=new -> RefreshLink; si falla, eliminar tabla)
$totalOk = 0
$totalBroken = 0
$brokenList = New-Object System.Collections.ArrayList

foreach ($backendPath in $copiedFiles) {
    $fileName = [System.IO.Path]::GetFileName($backendPath)
    $plan = $backendPlan[$fileName]
    if (@($plan).Count -eq 0) { continue }

    Write-Status "  Procesando: $fileName" -Color Cyan
    $db = $ada.OpenDatabase($backendPath, $false, $false, ";pwd=$backendPassword")

    foreach ($item in $plan) {
        # Backend inexistente en destino: eliminar link
        if (-not $item.BackendExists) {
            try { $db.TableDefs.Delete($item.TableName) } catch { }
            $totalBroken++
            [void]$brokenList.Add("$fileName::$($item.TableName) -> $($item.LinkedBackend) (backend no existe)")
            Write-Status "    BROKEN: $($item.TableName) -> $($item.LinkedBackend) (backend no existe)" -Color Yellow
            continue
        }

        $td = $db.TableDefs[$item.TableName]
        $td.Connect = "MS Access;PWD=$backendPassword;DATABASE=$($item.DestPath)"
        try {
            $td.RefreshLink()
            $totalOk++
            Write-Status "    $($item.TableName) -> $($item.LinkedBackend)" -Color Gray
        } catch {
            # RefreshLink fallo: eliminar la tabla. Nunca queda referencia a produccion.
            try { $db.TableDefs.Delete($item.TableName) } catch { }
            $totalBroken++
            [void]$brokenList.Add("$fileName::$($item.TableName) -> $($item.LinkedBackend) (RefreshLink fallo)")
            Write-Status "    BROKEN: $($item.TableName) -> $($item.LinkedBackend)" -Color Yellow
        }
    }
    $db.Close()
}

# === SUMMARY ===
Write-Status "=== Resumen ===" -Color Green
Write-Status "Vinculos OK:          $totalOk" -Color $(if ($totalOk -gt 0) { [ConsoleColor]::Green } else { [ConsoleColor]::Gray })
Write-Status "Vinculos eliminados:  $totalBroken" -Color $(if ($totalBroken -gt 0) { [ConsoleColor]::Yellow } else { [ConsoleColor]::Gray })
if (@($brokenList).Count -gt 0) {
    Write-Status "Detalles:" -Color Yellow
    foreach ($b in $brokenList) {
        Write-Status "  - $b" -Color Yellow
    }
}
Write-Status "Fin: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Color Gray

# Single GC at end
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ada) | Out-Null
[GC]::Collect()
