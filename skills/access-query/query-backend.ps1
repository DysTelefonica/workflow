param(
    [Parameter(Mandatory=$false)]
    [string]$SQL,
    [Parameter(Mandatory=$false)]
    [string]$Table,
    [Parameter(Mandatory=$false)]
    [string]$Field,
    [Parameter(Mandatory=$false)]
    [int]$Top,
    [Parameter(Mandatory=$false)]
    [switch]$Count,
    [Parameter(Mandatory=$false)]
    [switch]$Distinct,
    [Parameter(Mandatory=$false)]
    [switch]$ListTables,
    [Parameter(Mandatory=$false)]
    [switch]$LinkedTables,
    [Parameter(Mandatory=$false)]
    [switch]$GetSchema,
    [Parameter(Mandatory=$false)]
    [switch]$Compare,
    [Parameter(Mandatory=$false)]
    [string]$CompareBackend,
    [Parameter(Mandatory=$false)]
    [string]$CompareSQL,
    [Parameter(Mandatory=$false)]
    [string]$Backend = '',
    [Parameter(Mandatory=$false)]
    [string]$BackendPath = '',
    [Parameter(Mandatory=$false)]
    [string]$Password = ''
)

$ErrorActionPreference = 'Stop'
if ($PSScriptRoot) { $ScriptDir = $PSScriptRoot } else { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$configPath = Join-Path $ScriptDir 'backends.json'
if (-not (Test-Path $configPath)) {
    Write-Host 'ERROR: backends.json no encontrado' -ForegroundColor Red
    exit 1
}
$config = Get-Content $configPath | ConvertFrom-Json

$defaultBackend = $config.default

$backendMap = @{}
foreach ($prop in $config.backends.PSObject.Properties) {
    $backendMap[$prop.Name] = $prop.Value
}

function Get-Connection {
    param([string]$Path, [string]$Pw)
    $connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$Path;Jet OLEDB:Database Password=$Pw;"
    $conn = New-Object System.Data.OleDb.OleDbConnection($connString)
    $conn.Open()
    return $conn
}

function Get-BackendPath {
    param([string]$Name, [string]$OverridePath)
    if ($OverridePath) {
        if (-not (Test-Path $OverridePath)) {
            Write-Host 'ERROR: Ruta no encontrada' -ForegroundColor Red; exit 1
        }
        $pw = if ($Password) { $Password } else { '' }
        return @{ Path = $OverridePath; Password = $pw; Name = (Split-Path $OverridePath -Leaf) }
    }
    if ($Name -eq '') { $Name = $defaultBackend }
    if (-not $backendMap.ContainsKey($Name)) {
        Write-Host "ERROR: Backend '$Name' no encontrado. Disponibles: $($backendMap.Keys -join ', ')" -ForegroundColor Red; exit 1
    }
    $info = $backendMap[$Name]
    $pw = if ($Password) { $Password } else { $info.password }
    return @{ Path = $info.path; Password = $pw; Name = $Name }
}

function Format-Value {
    param($val)
    if ($val -is [System.DBNull] -or $null -eq $val) { return 'NULL' }
    if ($val -is [string] -and $val.Length -gt 50) { return $val.Substring(0, 47) + '...' }
    return [string]$val
}

# ── ListTables ────────────────────────────────────────────────────────────────
if ($ListTables) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $all = $conn.GetSchema('Tables')
    $conn.Close()
    Write-Host "=== TABLAS en $($t.Name) ===" -ForegroundColor Cyan
    foreach ($row in $all) {
        if ($row.Item('TABLE_TYPE') -eq 'TABLE') {
            Write-Host "  $($row.Item('TABLE_NAME'))" -ForegroundColor White
        }
    }
    exit 0
}

# ── LinkedTables ──────────────────────────────────────────────────────────────
if ($LinkedTables) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $all = $conn.GetSchema('Tables')
    $linked = @()
    foreach ($row in $all) {
        if ($row.Item('TABLE_TYPE') -eq 'LINK') {
            $linked += [PSCustomObject]@{ Name = $row.Item('TABLE_NAME'); Origin = $row.Item('TABLE_DESCRIPTION') }
        }
    }
    $conn.Close()
    Write-Host "=== TABLAS LINKED en $($t.Name) ===" -ForegroundColor Cyan
    if ($linked.Count -eq 0) { Write-Host '  (ninguna)' -ForegroundColor Gray }
    else { $linked | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Yellow; Write-Host "    -> $($_.Origin)" -ForegroundColor Gray } }
    exit 0
}

# ── GetSchema ─────────────────────────────────────────────────────────────────
if ($GetSchema -and $Table) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT * FROM [$Table] WHERE 1=0"
    try { $reader = $cmd.ExecuteReader() } catch { Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red; $conn.Close(); exit 1 }
    $dt = $reader.GetSchemaTable()
    $reader.Close(); $conn.Close()
    $colWidth = ($dt | ForEach-Object { $_.Item('ColumnName').Length } | Measure-Object -Maximum).Maximum
    if ($colWidth -lt 20) { $colWidth = 20 }
    $sep = '-' * $colWidth
    Write-Host "=== ESQUEMA: $Table ($($t.Name)) ===" -ForegroundColor Cyan
    Write-Host ''
    $header1 = "  | " + ("Campo".PadRight($colWidth)) + " | Tipo           | Nullable |"
    Write-Host $header1 -ForegroundColor White
    Write-Host "  | $sep | -------------- | -------- |" -ForegroundColor White
    foreach ($row in $dt) {
        $name     = $row.Item('ColumnName')
        $size     = $row.Item('ColumnSize')
        $nullable = if ($row.Item('AllowDBNull')) { 'Yes' } else { 'No' }
        $dtype    = $row.Item('DataTypeName')
        if     ($dtype -eq 'System.String')   { $tipo = "String($size)" }
        elseif ($dtype -eq 'System.Boolean')  { $tipo = 'Boolean' }
        elseif ($dtype -eq 'System.DateTime') { $tipo = 'Date' }
        elseif ($dtype -eq 'System.Double')   { $tipo = 'Double' }
        elseif ($dtype -eq 'System.Int64')    { $tipo = 'Long' }
        elseif ($dtype -eq 'System.Int32')    { $tipo = 'Integer' }
        else                                  { $tipo = $dtype }
        Write-Host "  | $($name.PadRight($colWidth)) | $($tipo.PadRight(14)) | $nullable        |" -ForegroundColor Green
    }
    exit 0
}

# ── Count ─────────────────────────────────────────────────────────────────────
if ($Count -and $Table) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) FROM [$Table]"
    try { $total = $cmd.ExecuteScalar() } catch { Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red; $conn.Close(); exit 1 }
    $conn.Close()
    Write-Host "=== COUNT: $Table ($($t.Name)) ===" -ForegroundColor Cyan
    Write-Host "  Total: $total" -ForegroundColor Green
    exit 0
}

# ── Distinct ──────────────────────────────────────────────────────────────────
if ($Distinct -and $Table -and $Field) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT DISTINCT [$Field] FROM [$Table] WHERE [$Field] IS NOT NULL ORDER BY [$Field]"
    try {
        $reader = $cmd.ExecuteReader()
        $vals = @()
        while ($reader.Read()) { $vals += Format-Value $reader[0] }
        $reader.Close()
    } catch { Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red; $conn.Close(); exit 1 }
    $conn.Close()
    Write-Host "=== DISTINCT $Field ON $Table ($($t.Name)) ===" -ForegroundColor Cyan
    Write-Host "  $($vals.Count) valores:" -ForegroundColor Gray
    $vals | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    exit 0
}

# ── Compare ───────────────────────────────────────────────────────────────────
if ($Compare -and $CompareSQL) {
    $leftName  = if ($Backend)        { $Backend }        else { $defaultBackend }
    $rightName = if ($CompareBackend) { $CompareBackend } else { $defaultBackend }
    $left  = Get-BackendPath -Name $leftName
    $right = Get-BackendPath -Name $rightName
    Write-Host '=== COMPARE ===' -ForegroundColor Cyan
    Write-Host "SQL   : $CompareSQL" -ForegroundColor White
    Write-Host "Left  : $($left.Name)" -ForegroundColor Yellow
    Write-Host "Right : $($right.Name)" -ForegroundColor Yellow
    Write-Host ''
    function Get-Ids {
        param($path, $pw, $q)
        $c = Get-Connection -Path $path -Pw $pw
        $cmd = $c.CreateCommand(); $cmd.CommandText = $q
        $r = $cmd.ExecuteReader(); $ids = @()
        while ($r.Read()) { $ids += $r[0] }
        $r.Close(); $c.Close()
        return ($ids | Sort-Object)
    }
    try {
        $idsL = Get-Ids -path $left.Path  -pw $left.Password  -q $CompareSQL
        $idsR = Get-Ids -path $right.Path -pw $right.Password -q $CompareSQL
    } catch { Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red; exit 1 }
    Write-Host "  $($left.Name)  : $($idsL.Count) filas" -ForegroundColor Yellow
    Write-Host "  $($right.Name): $($idsR.Count) filas" -ForegroundColor Yellow
    $onlyL = $idsL | Where-Object { $_ -notin $idsR }
    $onlyR = $idsR | Where-Object { $_ -notin $idsL }
    if ($onlyL.Count -eq 0 -and $onlyR.Count -eq 0) { Write-Host '  RESULT: IDENTICOS' -ForegroundColor Green }
    else {
        Write-Host '  RESULT: DIFERENTES' -ForegroundColor Red
        if ($onlyL.Count -gt 0) { Write-Host "  Solo en $($left.Name) ($($onlyL.Count)): $($onlyL -join ', ')" -ForegroundColor Yellow }
        if ($onlyR.Count -gt 0) { Write-Host "  Solo en $($right.Name) ($($onlyR.Count)): $($onlyR -join ', ')" -ForegroundColor Yellow }
    }
    exit 0
}

# ── SQL libre ─────────────────────────────────────────────────────────────────
if ($SQL) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    Write-Host "=== SQL ($($t.Name)) ===" -ForegroundColor Cyan
    Write-Host $SQL -ForegroundColor Gray
    Write-Host ''
    $maxRows = if ($Top -gt 0) { $Top } else { 20 }
    try {
        $conn = Get-Connection -Path $t.Path -Pw $t.Password
        $cmd = $conn.CreateCommand(); $cmd.CommandText = $SQL
        $reader = $cmd.ExecuteReader(); $rows = 0
        while ($reader.Read()) {
            $rows++
            if ($rows -le $maxRows) {
                $cols = @()
                for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                    $cols += "$($reader.GetName($i))=$(Format-Value $reader.GetValue($i))"
                }
                Write-Host "  [$rows] $($cols -join ' | ')" -ForegroundColor Green
            }
        }
        if ($rows -gt $maxRows) { Write-Host "  ... y $($rows - $maxRows) filas mas (limitado a $maxRows)" -ForegroundColor Gray }
        Write-Host ''
        Write-Host "  Total: $rows filas" -ForegroundColor Cyan
        $reader.Close(); $conn.Close()
    } catch { Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red; exit 1 }
    exit 0
}

# ── Help ──────────────────────────────────────────────────────────────────────
Write-Host @"
ACCESS-QUERY — Consultas SQL a backends Access (.accdb)
Config: backends.json en el mismo directorio que el script.

MODO LIBRE:
  -SQL "SELECT ..."            Ejecutar SQL libre
  -SQL "SELECT ..." -Top 5    Con limite de filas
  -BackendPath "ruta.accdb"   Ruta directa (ignora backends.json)

MODO INSPECCION:
  -GetSchema -Table TbUsuarios      Ver campos de tabla
  -Count -Table TbUsuarios         Contar registros
  -Distinct -Table Tb -Field C     Valores unicos
  -ListTables                       Listar tablas locales
  -LinkedTables                     Listar tablas linked

MODO COMPARAR:
  -Compare -CompareSQL "SELECT ..." -Backend A -CompareBackend B

PARAMETROS:
  -Backend          Nombre de backend (definido en backends.json)
  -BackendPath      Ruta directa al .accdb
  -Password         Password override del .accdb
  -Top              Limite de filas (default: 20)

BACKENDS CONFIGURADOS: $($backendMap.Keys -join ', ')
Default: $defaultBackend
"@ -ForegroundColor White
