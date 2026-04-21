param(
    # -- Modos de lectura --
    [Parameter(Mandatory=$false)] [string]$SQL            = '',
    [Parameter(Mandatory=$false)] [string]$Table          = '',
    [Parameter(Mandatory=$false)] [string]$Field          = '',
    [Parameter(Mandatory=$false)] [int]$Top               = -1,   # -1 = usar default 20; 0 = sin límite
    [Parameter(Mandatory=$false)] [switch]$Count,
    [Parameter(Mandatory=$false)] [switch]$Distinct,
    [Parameter(Mandatory=$false)] [switch]$ListTables,
    [Parameter(Mandatory=$false)] [switch]$LinkedTables,
    [Parameter(Mandatory=$false)] [switch]$GetSchema,
    [Parameter(Mandatory=$false)] [switch]$Compare,
    [Parameter(Mandatory=$false)] [string]$CompareBackend = '',
    [Parameter(Mandatory=$false)] [string]$CompareSQL     = '',

    # -- Modos de escritura --
    [Parameter(Mandatory=$false)] [string]$Exec           = '',
    [Parameter(Mandatory=$false)] [string]$Script         = '',
    [Parameter(Mandatory=$false)] [switch]$Seed,
    [Parameter(Mandatory=$false)] [switch]$Teardown,
    [Parameter(Mandatory=$false)] [string]$FixtureTag     = '',
    [Parameter(Mandatory=$false)] [switch]$CreateTable,
    [Parameter(Mandatory=$false)] [switch]$DropTable,

    # -- Guardas de seguridad --
    [Parameter(Mandatory=$false)] [switch]$DryRun,
    [Parameter(Mandatory=$false)] [string]$AllowTable     = '',
    [Parameter(Mandatory=$false)] [string]$DenyTable      = '',
    [Parameter(Mandatory=$false)] [switch]$StrictWrite,
    [Parameter(Mandatory=$false)] [switch]$Force,

    # -- Salida --
    [Parameter(Mandatory=$false)] [switch]$Json,

    # -- Conexion --
    [Parameter(Mandatory=$false)] [string]$Backend        = '',
    [Parameter(Mandatory=$false)] [string]$BackendPath    = '',
    [Parameter(Mandatory=$false)] [string]$Password       = '__UNSET__'   # sentinel para distinguir "" de no-pasado
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# VALIDACION DE COMBINACIONES INCOMPATIBLES
# =============================================================================

$validationErrors = @()

$sqlModeProvided      = $SQL -ne ''
$execModeProvided     = $Exec -ne ''
$scriptModeProvided   = $Script -ne ''
$compareModeProvided  = $CompareSQL -ne ''
$listTablesProvided   = $ListTables
$linkedTablesProvided = $LinkedTables
$getSchemaProvided    = $GetSchema
$countProvided        = $Count
$distinctProvided     = $Distinct
$createTableProvided  = $CreateTable
$dropTableProvided    = $DropTable

$readModes = @($getSchemaProvided, $countProvided, $distinctProvided, $compareModeProvided,
                $listTablesProvided, $linkedTablesProvided)
$activeReadModes = ($readModes | Where-Object { $_ }).Count

if ($activeReadModes -gt 1) {
    $validationErrors += "Solo puedes usar un modo de lectura a la vez (-GetSchema, -Count, -Distinct, -Compare, -ListTables, -LinkedTables)."
}
if ($execModeProvided -and $scriptModeProvided) {
    $validationErrors += "-Exec y -Script son mutuamente excluyentes (elige uno)."
}
if ($sqlModeProvided -and ($execModeProvided -or $scriptModeProvided)) {
    $validationErrors += "-SQL es modo de solo lectura: no puede combinarse con -Exec o -Script."
}
if ($sqlModeProvided -and $compareModeProvided) {
    $validationErrors += "-SQL y -Compare son mutuamente excluyentes."
}
if ($sqlModeProvided -and ($getSchemaProvided -or $countProvided -or $distinctProvided -or $listTablesProvided -or $linkedTablesProvided)) {
    $validationErrors += "-SQL no puede combinarse con otros modos de lectura."
}
if ($getSchemaProvided -and $execModeProvided) {
    $validationErrors += "-GetSchema es modo de solo lectura: no puede combinarse con -Exec."
}
if ($Compare -and ($execModeProvided -or $scriptModeProvided -or $Seed -or $Teardown)) {
    $validationErrors += "-Compare es modo de solo lectura."
}
if ($Seed -and $Teardown) { $validationErrors += "-Seed y -Teardown son mutuamente excluyentes." }
if (($Seed -or $Teardown) -and ($createTableProvided -or $dropTableProvided)) {
    $validationErrors += "-Seed/-Teardown no pueden combinarse con DDL (-CreateTable/-DropTable)."
}
if ($createTableProvided -and $dropTableProvided) {
    $validationErrors += "-DropTable y -CreateTable son mutuamente excluyentes."
}

if ($validationErrors.Count -gt 0) {
    foreach ($err in $validationErrors) { Write-Host "ERROR: $err" -ForegroundColor Red }
    Write-Host "Ejecuta '.\query-backend.ps1' sin argumentos para ver la ayuda." -ForegroundColor Yellow
    exit 1
}

# =============================================================================
# DISPATCHER
# =============================================================================

$Mode = $null
$SqlInput = $null

if     ($Seed -and $Exec)            { $Mode = 'Seed';      $SqlInput = $Exec }
elseif ($Seed -and $Script)          { $Mode = 'Seed';      $SqlInput = '__FILE__' }
elseif ($Teardown -and $Exec)        { $Mode = 'Teardown';  $SqlInput = $Exec }
elseif ($Teardown -and $Script)      { $Mode = 'Teardown';  $SqlInput = '__FILE__' }
elseif ($CreateTable -and $Exec)     { $Mode = 'DDL';       $SqlInput = $Exec }
elseif ($DropTable -and $Table)      { $Mode = 'DDL';       $SqlInput = "DROP TABLE [$Table]" }
elseif ($Exec)                       { $Mode = 'Exec';      $SqlInput = $Exec }
elseif ($Script)                     { $Mode = 'Script';    $SqlInput = '__FILE__' }
elseif ($ListTables)                 { $Mode = 'ListTables' }
elseif ($LinkedTables)               { $Mode = 'LinkedTables' }
elseif ($GetSchema -and $Table)      { $Mode = 'GetSchema' }
elseif ($Count -and $Table)          { $Mode = 'Count' }
elseif ($Distinct -and $Table -and $Field) { $Mode = 'Distinct' }
elseif ($Compare -and $CompareSQL)   { $Mode = 'Compare' }
elseif ($SQL)                        { $Mode = 'SQL' }
elseif ($Seed)       { Write-Host 'ERROR: -Seed requiere -Exec "SQL" o -Script "ruta.sql"' -ForegroundColor Red; exit 1 }
elseif ($Teardown)   { Write-Host 'ERROR: -Teardown requiere -Exec "SQL" o -Script "ruta.sql"' -ForegroundColor Red; exit 1 }
elseif ($GetSchema)  { Write-Host 'ERROR: -GetSchema requiere -Table "nombre"' -ForegroundColor Red; exit 1 }
elseif ($Count)      { Write-Host 'ERROR: -Count requiere -Table "nombre"' -ForegroundColor Red; exit 1 }
elseif ($Distinct)   { Write-Host 'ERROR: -Distinct requiere -Table "nombre" -Field "campo"' -ForegroundColor Red; exit 1 }
elseif ($Compare)    { Write-Host 'ERROR: -Compare requiere -CompareSQL "SELECT ..."' -ForegroundColor Red; exit 1 }
elseif ($DropTable)  { Write-Host 'ERROR: -DropTable requiere -Table "nombre"' -ForegroundColor Red; exit 1 }
elseif ($CreateTable){ Write-Host 'ERROR: -CreateTable requiere -Exec "CREATE TABLE ..."' -ForegroundColor Red; exit 1 }

# =============================================================================
# CONFIGURACION
# =============================================================================

if ($PSScriptRoot) { $ScriptDir = $PSScriptRoot } else { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$configPath = Join-Path $ScriptDir 'backends.json'
if (-not (Test-Path $configPath)) {
    Write-Host 'ERROR: backends.json no encontrado' -ForegroundColor Red
    exit 1
}
$config = Get-Content $configPath -Raw | ConvertFrom-Json
$defaultBackend = $config.default

$backendMap = @{}
foreach ($prop in $config.backends.PSObject.Properties) {
    $backendMap[$prop.Name] = $prop.Value
}

# -- Deny-list global --
$globalDenyTables = @()
if ($config.PSObject.Properties['deny_tables']) { $globalDenyTables = @($config.deny_tables) }
if ($DenyTable) { $globalDenyTables += ($DenyTable -split ',') | ForEach-Object { $_.Trim() } }
$globalDenyTables = $globalDenyTables | Select-Object -Unique

# -- Allow-list --
$allowTableList = @()
if ($AllowTable) { $allowTableList = ($AllowTable -split ',') | ForEach-Object { $_.Trim() } }

# =============================================================================
# RESOLUCION DE PASSWORDS
# Cadena de prioridad:
#   1. -Password (CLI override)  — si es '' explícito, se usa '' (sin password)
#   2. Env var ACCESS_QUERY_PW_<BACKEND>
#   3. Env var ACCESS_QUERY_PASSWORD (global)
#   4. .secrets.json
#   5. backends.json > password (DEPRECADO)
#   6. '' — sin password (BDs sin contraseña son comunes en desarrollo)
# =============================================================================

function Resolve-Password {
    param([string]$BackendName, [string]$CliPassword, [bool]$CliPasswordSet, [string]$JsonPassword)

    # 1. CLI override explícito (incluyendo string vacío = sin password)
    if ($CliPasswordSet) { return $CliPassword }

    # 2. Env var por backend
    $envPerBackend = "ACCESS_QUERY_PW_$($BackendName -replace '[^a-zA-Z0-9]','_')"
    $envVal = [System.Environment]::GetEnvironmentVariable($envPerBackend)
    if ($null -ne $envVal -and $envVal -ne '') { return $envVal }

    # 3. Env var global
    $envGlobal = [System.Environment]::GetEnvironmentVariable('ACCESS_QUERY_PASSWORD')
    if ($null -ne $envGlobal -and $envGlobal -ne '') { return $envGlobal }

    # 4. .secrets.json
    $secretsPath = Join-Path $ScriptDir '.secrets.json'
    if (Test-Path $secretsPath) {
        try {
            $secrets = Get-Content $secretsPath -Raw | ConvertFrom-Json
            if ($secrets.PSObject.Properties[$BackendName]) { return $secrets.$BackendName }
            if ($secrets.PSObject.Properties['default']) { return $secrets.default }
        } catch { }
    }

    # 5. backends.json (backward compat)
    if ($JsonPassword -and $JsonPassword -ne '') { return $JsonPassword }

    # 6. Sin password (BDs de desarrollo sin contraseña)
    return ''
}

$cliPasswordSet = ($Password -ne '__UNSET__')
$cliPasswordValue = if ($cliPasswordSet) { $Password } else { '' }

# =============================================================================
# FUNCIONES UTILITARIAS
# =============================================================================

function Get-Connection {
    param([string]$Path, [string]$Pw)
    if ($Pw -ne '') {
        $connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$Path;Jet OLEDB:Database Password=$Pw;"
    } else {
        $connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$Path;"
    }
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
        $resolvedPw = Resolve-Password -BackendName '__direct__' -CliPassword $cliPasswordValue -CliPasswordSet $cliPasswordSet -JsonPassword ''
        return @{ Path = $OverridePath; Password = $resolvedPw; Name = (Split-Path $OverridePath -Leaf) }
    }
    if ($Name -eq '') { $Name = $defaultBackend }
    if (-not $backendMap.ContainsKey($Name)) {
        Write-Host "ERROR: Backend '$Name' no encontrado. Disponibles: $($backendMap.Keys -join ', ')" -ForegroundColor Red; exit 1
    }
    $info = $backendMap[$Name]
    $jsonPw = if ($info.PSObject.Properties['password']) { $info.password } else { '' }
    $resolvedPw = Resolve-Password -BackendName $Name -CliPassword $cliPasswordValue -CliPasswordSet $cliPasswordSet -JsonPassword $jsonPw
    return @{ Path = $info.path; Password = $resolvedPw; Name = $Name }
}

function Format-Value {
    param($val)
    if ($val -is [System.DBNull] -or $null -eq $val) { return $null }   # null real para JSON
    if ($val -is [string] -and $val.Length -gt 200) { return $val.Substring(0, 197) + '...' }
    if ($val -is [DateTime]) { return $val.ToString('yyyy-MM-dd HH:mm:ss') }
    return $val
}

function Format-ValueDisplay {
    param($val)
    $v = Format-Value $val
    if ($null -eq $v) { return 'NULL' }
    if ($v -is [string] -and $v.Length -gt 80) { return $v.Substring(0, 77) + '...' }
    return [string]$v
}

# =============================================================================
# PARSER DE SENTENCIAS SQL
# State machine que respeta ; dentro de strings (' y ") y comentarios --
# =============================================================================

function Split-SqlStatements {
    param([string]$SqlBlock)
    $statements = [System.Collections.Generic.List[string]]::new()
    $current = [System.Text.StringBuilder]::new()
    $inSingleQuote = $false
    $inDoubleQuote = $false

    for ($i = 0; $i -lt $SqlBlock.Length; $i++) {
        $c = $SqlBlock[$i]

        if ($c -eq "'" -and -not $inDoubleQuote) {
            if ($inSingleQuote -and ($i + 1) -lt $SqlBlock.Length -and $SqlBlock[$i + 1] -eq "'") {
                [void]$current.Append($c); $i++; [void]$current.Append($SqlBlock[$i]); continue
            }
            $inSingleQuote = -not $inSingleQuote; [void]$current.Append($c); continue
        }
        if ($c -eq '"' -and -not $inSingleQuote) {
            $inDoubleQuote = -not $inDoubleQuote; [void]$current.Append($c); continue
        }
        if ($c -eq '-' -and -not $inSingleQuote -and -not $inDoubleQuote) {
            if (($i + 1) -lt $SqlBlock.Length -and $SqlBlock[$i + 1] -eq '-') {
                while ($i -lt $SqlBlock.Length -and $SqlBlock[$i] -ne "`n") { $i++ }
                continue
            }
        }
        if ($c -eq ';' -and -not $inSingleQuote -and -not $inDoubleQuote) {
            $stmt = $current.ToString().Trim()
            if ($stmt) { [void]$statements.Add($stmt) }
            [void]$current.Clear(); continue
        }
        [void]$current.Append($c)
    }
    $lastStmt = $current.ToString().Trim()
    if ($lastStmt) { [void]$statements.Add($lastStmt) }
    return $statements.ToArray()
}

# =============================================================================
# SEGURIDAD
# =============================================================================

function Get-LinkedTableNames {
    param([System.Data.OleDb.OleDbConnection]$Conn)
    $linked = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $schema = $Conn.GetSchema('Tables')
    foreach ($row in $schema) {
        if ($row.Item('TABLE_TYPE') -eq 'LINK') { [void]$linked.Add($row.Item('TABLE_NAME')) }
    }
    return $linked
}

function Extract-TableNames {
    param([string]$SqlText, [switch]$WriteTargetsOnly)
    $tables = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $writePatterns = @(
        'INSERT\s+INTO\s+\[?([^\]\s\(]+)\]?',
        'UPDATE\s+\[?([^\]\s]+)\]?',
        'DELETE\s+FROM\s+\[?([^\]\s]+)\]?',
        'DROP\s+TABLE\s+\[?([^\]\s]+)\]?',
        'CREATE\s+TABLE\s+\[?([^\]\s]+)\]?',
        'ALTER\s+TABLE\s+\[?([^\]\s]+)\]?',
        'TRUNCATE\s+TABLE\s+\[?([^\]\s]+)\]?'
    )
    foreach ($pattern in $writePatterns) {
        $matches_ = [regex]::Matches($SqlText, $pattern, 'IgnoreCase')
        foreach ($m in $matches_) { [void]$tables.Add($m.Groups[1].Value) }
    }
    if (-not $WriteTargetsOnly) {
        $readPatterns = @('FROM\s+\[?([^\]\s,\(]+)\]?', 'JOIN\s+\[?([^\]\s]+)\]?')
        foreach ($pattern in $readPatterns) {
            $matches_ = [regex]::Matches($SqlText, $pattern, 'IgnoreCase')
            foreach ($m in $matches_) { [void]$tables.Add($m.Groups[1].Value) }
        }
    }
    return $tables
}

function Test-IsWriteStatement {
    param([string]$SqlText)
    return ($SqlText.TrimStart() -match '^\s*(INSERT|UPDATE|DELETE|DROP|CREATE|ALTER|TRUNCATE)\s')
}

function Assert-WriteAllowed {
    param(
        [string]$SqlText,
        [System.Collections.Generic.HashSet[string]]$LinkedTables,
        [string[]]$DenyList,
        [string[]]$AllowList
    )
    $targets = Extract-TableNames -SqlText $SqlText -WriteTargetsOnly
    $blocked = @()
    foreach ($tbl in $targets) {
        foreach ($deny in $DenyList) {
            if ($tbl -like $deny) { $blocked += "DENY: '$tbl' coincide con pattern '$deny'" }
        }
        if ($LinkedTables.Contains($tbl)) { $blocked += "LINKED: '$tbl' es tabla LINKED/EXTERNA" }
        if ($AllowList.Count -gt 0) {
            $allowed = $false
            foreach ($allow in $AllowList) { if ($tbl -like $allow) { $allowed = $true; break } }
            if (-not $allowed) { $blocked += "ALLOW: '$tbl' no esta en allow-list ($($AllowList -join ', '))" }
        }
    }
    return @{ Targets = $targets; Blocked = $blocked }
}

# =============================================================================
# VALIDACION ESTRICTA SEED/TEARDOWN
# =============================================================================

$isWriteMode = $Mode -in @('Exec','Script','Seed','Teardown','DDL')

if ($isWriteMode) {
    if ($Mode -in @('Seed','Teardown') -and $allowTableList.Count -eq 0 -and -not $Force) {
        Write-Host "ERROR: -$Mode requiere -AllowTable para evitar tocar tablas equivocadas." -ForegroundColor Red
        Write-Host '  Ejemplo: -AllowTable "TbSolicitudes,TbDocumentos"' -ForegroundColor Yellow
        Write-Host '  Usa -Force para saltarte esta restriccion.' -ForegroundColor DarkYellow
        exit 1
    }
    if ($StrictWrite -and $allowTableList.Count -eq 0 -and -not $Force) {
        Write-Host 'ERROR: -StrictWrite activo: se requiere -AllowTable explicito.' -ForegroundColor Red
        Write-Host '  Ejemplo: -AllowTable "TbSolicitudes"' -ForegroundColor Yellow
        exit 1
    }
}

# =============================================================================
# MOTOR DE EJECUCION CON FIXTURE TRACKING + SALIDA JSON
# =============================================================================

function Invoke-WriteStatements {
    param(
        [string[]]$Statements,
        [hashtable]$BackendInfo,
        [switch]$IsDryRun,
        [string]$Label = 'EXEC',
        [string]$FixtureTagValue = '',
        [switch]$OutputJson
    )

    $conn = Get-Connection -Path $BackendInfo.Path -Pw $BackendInfo.Password
    $linkedSet = Get-LinkedTableNames -Conn $conn

    $report = @{
        mode          = $Label
        backend       = $BackendInfo.Name
        backendPath   = $BackendInfo.Path
        dryRun        = [bool]$IsDryRun
        fixtureTag    = $FixtureTagValue
        timestamp     = (Get-Date -Format 'o')
        denyList      = $globalDenyTables
        allowList     = $allowTableList
        linkedCount   = $linkedSet.Count
        statements    = @()
        aborted       = $false
        totalAffected = 0
        tablesWritten = @()
    }
    $allTablesWritten = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if (-not $OutputJson) {
        Write-Host "=== $Label ($($BackendInfo.Name)) ===" -ForegroundColor Cyan
        if ($FixtureTagValue) { Write-Host "  Fixture Tag: $FixtureTagValue" -ForegroundColor Magenta }
        if ($globalDenyTables.Count -gt 0) { Write-Host "  Deny-list: $($globalDenyTables -join ', ')" -ForegroundColor DarkYellow }
        if ($allowTableList.Count -gt 0) { Write-Host "  Allow-list: $($allowTableList -join ', ')" -ForegroundColor DarkGreen }
        Write-Host "  Linked tables: $($linkedSet.Count)" -ForegroundColor Gray
        if ($IsDryRun) { Write-Host "  ** MODO DRY-RUN **" -ForegroundColor Magenta }
        Write-Host ''
    }

    $stmtIndex = 0
    foreach ($stmt in $Statements) {
        $stmtIndex++
        $isWrite = Test-IsWriteStatement -SqlText $stmt
        $stmtResult = @{
            index    = $stmtIndex
            sql      = $stmt
            type     = if ($isWrite) { 'WRITE' } else { 'READ' }
            status   = 'PENDING'
            targets  = @()
            blocked  = @()
            affected = 0
        }

        if ($isWrite) {
            $check = Assert-WriteAllowed -SqlText $stmt -LinkedTables $linkedSet `
                -DenyList $globalDenyTables -AllowList $allowTableList
            $stmtResult.targets = @($check.Targets)
            foreach ($tbl in $check.Targets) { [void]$allTablesWritten.Add($tbl) }

            if ($check.Blocked.Count -gt 0) {
                $stmtResult.status = 'BLOCKED'
                $stmtResult.blocked = $check.Blocked
                $report.statements += $stmtResult
                $report.aborted = $true
                if (-not $OutputJson) {
                    Write-Host "  [$stmtIndex] BLOQUEADO" -ForegroundColor Red
                    foreach ($b in $check.Blocked) { Write-Host "      $b" -ForegroundColor Red }
                    $dSql = if ($stmt.Length -gt 100) { $stmt.Substring(0, 97) + '...' } else { $stmt }
                    Write-Host "      SQL: $dSql" -ForegroundColor DarkGray
                    Write-Host '  ABORTADO por guarda de seguridad.' -ForegroundColor Red
                }
                break
            }
            if (-not $OutputJson) {
                $tl = ($check.Targets | ForEach-Object { $_ }) -join ', '
                Write-Host "  [$stmtIndex] WRITE -> $tl" -ForegroundColor Yellow
            }
        } else {
            if (-not $OutputJson) { Write-Host "  [$stmtIndex] READ" -ForegroundColor Green }
        }

        $dSql = if ($stmt.Length -gt 120) { $stmt.Substring(0, 117) + '...' } else { $stmt }
        if (-not $OutputJson) { Write-Host "      $dSql" -ForegroundColor Gray }

        if ($IsDryRun) {
            $stmtResult.status = 'DRY-RUN'
            $report.statements += $stmtResult
            if (-not $OutputJson) { Write-Host "      [DRY-RUN] No ejecutado" -ForegroundColor Magenta }
            continue
        }

        try {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $stmt
            if ($isWrite) {
                $affected = $cmd.ExecuteNonQuery()
                $stmtResult.affected = $affected
                $stmtResult.status = 'OK'
                $report.totalAffected += $affected
                if (-not $OutputJson) { Write-Host "      OK: $affected filas" -ForegroundColor Green }
            } else {
                $reader = $cmd.ExecuteReader()
                $rows = 0; $maxShow = 10
                while ($reader.Read()) {
                    $rows++
                    if ($rows -le $maxShow -and -not $OutputJson) {
                        $cols = @()
                        for ($ci = 0; $ci -lt $reader.FieldCount; $ci++) {
                            $cols += "$($reader.GetName($ci))=$(Format-ValueDisplay $reader.GetValue($ci))"
                        }
                        Write-Host "      [$rows] $($cols -join ' | ')" -ForegroundColor Green
                    }
                }
                $reader.Close()
                $stmtResult.affected = $rows
                $stmtResult.status = 'OK'
                if (-not $OutputJson) {
                    if ($rows -gt $maxShow) { Write-Host "      ... +$($rows - $maxShow) filas" -ForegroundColor Gray }
                    Write-Host "      Total: $rows filas" -ForegroundColor Cyan
                }
            }
        } catch {
            $stmtResult.status = 'ERROR'
            $stmtResult.blocked = @($_.Exception.Message)
            $report.statements += $stmtResult
            $report.aborted = $true
            if (-not $OutputJson) {
                Write-Host "      ERROR: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host '  ABORTADO por error SQL.' -ForegroundColor Red
            }
            break
        }
        $report.statements += $stmtResult
    }

    $report.tablesWritten = @($allTablesWritten)
    $conn.Close()

    if (-not $OutputJson) {
        Write-Host ''
        Write-Host '  -- RESUMEN --' -ForegroundColor Cyan
        if ($FixtureTagValue) { Write-Host "  Fixture: $FixtureTagValue" -ForegroundColor Magenta }
        Write-Host "  Sentencias: $($Statements.Count)" -ForegroundColor White
        $okC = ($report.statements | Where-Object { $_.status -eq 'OK' }).Count
        $blC = ($report.statements | Where-Object { $_.status -eq 'BLOCKED' }).Count
        $erC = ($report.statements | Where-Object { $_.status -eq 'ERROR' }).Count
        $drC = ($report.statements | Where-Object { $_.status -eq 'DRY-RUN' }).Count
        Write-Host "  OK: $okC | Blocked: $blC | Error: $erC | DryRun: $drC" -ForegroundColor White
        if ($allTablesWritten.Count -gt 0) { Write-Host "  Tablas: $($allTablesWritten -join ', ')" -ForegroundColor Yellow }
        if (-not $IsDryRun) { Write-Host "  Filas afectadas: $($report.totalAffected)" -ForegroundColor Green }
    }

    # Fixture log
    if ($FixtureTagValue -and -not $IsDryRun -and -not $report.aborted) {
        $logPath = Join-Path $ScriptDir '.fixture-log.json'
        $logEntry = @{
            fixtureTag = $FixtureTagValue
            mode       = $Label
            timestamp  = $report.timestamp
            backend    = $BackendInfo.Name
            allowList  = $allowTableList
            dryRun     = [bool]$IsDryRun
            aborted    = $report.aborted
            tables     = @($allTablesWritten)
            affected   = $report.totalAffected
            stmtCount  = $Statements.Count
        }
        $existingLog = @()
        if (Test-Path $logPath) {
            try { $existingLog = @(Get-Content $logPath -Raw | ConvertFrom-Json) } catch { $existingLog = @() }
        }
        $existingLog += $logEntry
        $existingLog | ConvertTo-Json -Depth 5 | Set-Content $logPath -Encoding UTF8
        if (-not $OutputJson) { Write-Host "  Fixture log: $logPath" -ForegroundColor DarkGray }
    }

    return $report
}

# =============================================================================
# EJECUCION POR MODO
# =============================================================================

# ---- ListTables ----
if ($Mode -eq 'ListTables') {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $all = $conn.GetSchema('Tables')
    $conn.Close()
    $tableNames = @($all | Where-Object { $_.Item('TABLE_TYPE') -eq 'TABLE' } | ForEach-Object { $_.Item('TABLE_NAME') })
    if ($Json) {
        @{ mode = 'ListTables'; backend = $t.Name; tables = $tableNames } | ConvertTo-Json -Depth 3 | Write-Output
    } else {
        Write-Host "=== TABLAS en $($t.Name) ($($tableNames.Count)) ===" -ForegroundColor Cyan
        foreach ($tn in $tableNames) { Write-Host "  $tn" -ForegroundColor White }
    }
    exit 0
}

# ---- LinkedTables ----
if ($Mode -eq 'LinkedTables') {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $all = $conn.GetSchema('Tables')
    $linked = @($all | Where-Object { $_.Item('TABLE_TYPE') -eq 'LINK' } | ForEach-Object {
        @{ name = $_.Item('TABLE_NAME'); origin = $_.Item('TABLE_DESCRIPTION') }
    })
    $conn.Close()
    if ($Json) {
        @{ mode = 'LinkedTables'; backend = $t.Name; tables = $linked } | ConvertTo-Json -Depth 3 | Write-Output
    } else {
        Write-Host "=== TABLAS LINKED en $($t.Name) ($($linked.Count)) ===" -ForegroundColor Cyan
        if ($linked.Count -eq 0) { Write-Host '  (ninguna)' -ForegroundColor Gray }
        else { $linked | ForEach-Object { Write-Host "  $($_.name)" -ForegroundColor Yellow; Write-Host "    -> $($_.origin)" -ForegroundColor Gray } }
    }
    exit 0
}

# ---- GetSchema ----
if ($Mode -eq 'GetSchema') {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT * FROM [$Table] WHERE 1=0"
    try { $reader = $cmd.ExecuteReader() } catch { Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red; $conn.Close(); exit 1 }
    $dt = $reader.GetSchemaTable()
    $reader.Close(); $conn.Close()

    $columns = @()
    foreach ($row in $dt) {
        $name     = $row.Item('ColumnName')
        $size     = $row.Item('ColumnSize')
        $nullable = $row.Item('AllowDBNull')
        $dtype    = $row.Item('DataTypeName')
        $tipo = switch -Wildcard ($dtype) {
            'System.String'   { "String($size)" }
            'System.Boolean'  { 'Boolean' }
            'System.DateTime' { 'Date' }
            'System.Double'   { 'Double' }
            'System.Int64'    { 'Long' }
            'System.Int32'    { 'Integer' }
            'System.Decimal'  { 'Decimal' }
            default           { $dtype }
        }
        $columns += @{ name = $name; type = $tipo; nullable = $nullable }
    }

    if ($Json) {
        @{ mode = 'GetSchema'; table = $Table; backend = $t.Name; columns = $columns } | ConvertTo-Json -Depth 4 | Write-Output
    } else {
        $colWidth = [Math]::Max(20, ($columns | ForEach-Object { $_.name.Length } | Measure-Object -Maximum).Maximum)
        $sep = '-' * $colWidth
        Write-Host "=== ESQUEMA: $Table ($($t.Name)) ===" -ForegroundColor Cyan
        Write-Host ''
        Write-Host "  | $("Campo".PadRight($colWidth)) | Tipo           | Nullable |" -ForegroundColor White
        Write-Host "  | $sep | -------------- | -------- |" -ForegroundColor White
        foreach ($col in $columns) {
            Write-Host "  | $($col.name.PadRight($colWidth)) | $($col.type.PadRight(14)) | $(if($col.nullable){'Yes'}else{'No'})        |" -ForegroundColor Green
        }
    }
    exit 0
}

# ---- Count ----
if ($Mode -eq 'Count') {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    $conn = Get-Connection -Path $t.Path -Pw $t.Password
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) FROM [$Table]"
    try { $total = $cmd.ExecuteScalar() } catch { Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red; $conn.Close(); exit 1 }
    $conn.Close()
    if ($Json) {
        @{ mode = 'Count'; table = $Table; backend = $t.Name; count = $total } | ConvertTo-Json | Write-Output
    } else {
        Write-Host "=== COUNT: $Table ($($t.Name)) ===" -ForegroundColor Cyan
        Write-Host "  Total: $total" -ForegroundColor Green
    }
    exit 0
}

# ---- Distinct ----
if ($Mode -eq 'Distinct') {
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
    if ($Json) {
        @{ mode = 'Distinct'; table = $Table; field = $Field; backend = $t.Name; count = $vals.Count; values = $vals } | ConvertTo-Json -Depth 3 | Write-Output
    } else {
        Write-Host "=== DISTINCT $Field ON $Table ($($t.Name)) ===" -ForegroundColor Cyan
        Write-Host "  $($vals.Count) valores:" -ForegroundColor Gray
        $vals | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    }
    exit 0
}

# ---- Compare ----
if ($Mode -eq 'Compare') {
    $leftName  = if ($Backend)        { $Backend }        else { $defaultBackend }
    $rightName = if ($CompareBackend) { $CompareBackend } else { $defaultBackend }
    $left  = Get-BackendPath -Name $leftName
    $right = Get-BackendPath -Name $rightName
    Write-Host '=== COMPARE ===' -ForegroundColor Cyan
    Write-Host "SQL   : $CompareSQL" -ForegroundColor White
    Write-Host "Left  : $($left.Name)" -ForegroundColor Yellow
    Write-Host "Right : $($right.Name)" -ForegroundColor Yellow
    Write-Host ''
    function Get-Ids { param($path, $pw, $q)
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

# ---- SQL (SELECT libre) ----
if ($Mode -eq 'SQL') {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath
    # -Top -1 → default 20; -Top 0 → sin límite; -Top N → N filas
    $maxRows = if ($Top -eq 0) { [int]::MaxValue } elseif ($Top -gt 0) { $Top } else { 20 }
    $unlimited = ($Top -eq 0)

    try {
        $conn = Get-Connection -Path $t.Path -Pw $t.Password
        $cmd = $conn.CreateCommand(); $cmd.CommandText = $SQL
        $reader = $cmd.ExecuteReader()

        # Leer columnas
        $colNames = @()
        for ($i = 0; $i -lt $reader.FieldCount; $i++) { $colNames += $reader.GetName($i) }

        $rows = @(); $rowCount = 0
        while ($reader.Read()) {
            $rowCount++
            if ($rowCount -le $maxRows) {
                $rowObj = @{}
                for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                    $rowObj[$colNames[$i]] = Format-Value $reader.GetValue($i)
                }
                $rows += $rowObj
            }
        }
        $reader.Close(); $conn.Close()
    } catch { Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red; exit 1 }

    $truncated = ($rowCount -gt $maxRows)

    if ($Json) {
        $out = @{
            mode     = 'SQL'
            backend  = $t.Name
            sql      = $SQL
            rowCount = $rowCount
            columns  = $colNames
            rows     = $rows
        }
        if ($truncated) { $out['truncated'] = $true; $out['shownRows'] = $maxRows }
        $out | ConvertTo-Json -Depth 5 | Write-Output
    } else {
        Write-Host "=== SQL ($($t.Name)) ===" -ForegroundColor Cyan
        Write-Host $SQL -ForegroundColor Gray
        Write-Host ''
        $r = 0
        foreach ($row in $rows) {
            $r++
            $cols = $colNames | ForEach-Object { "$_=$(if($null -eq $row[$_]){'NULL'}else{[string]$row[$_]})" }
            Write-Host "  [$r] $($cols -join ' | ')" -ForegroundColor Green
        }
        if ($truncated) { Write-Host "  ... y $($rowCount - $maxRows) filas mas (usa -Top 0 para todas)" -ForegroundColor Gray }
        Write-Host ''
        Write-Host "  Total: $rowCount filas$(if(-not $unlimited -and $Top -le 0){' (limitado a 20, usa -Top 0 para ver todas)'})" -ForegroundColor Cyan
    }
    exit 0
}

# ---- Modos de escritura ----
if ($Mode -in @('Exec','Script','Seed','Teardown','DDL')) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath

    if ($SqlInput -eq '__FILE__') {
        $scriptPath = $Script
        if (-not (Test-Path $scriptPath)) {
            Write-Host "ERROR: Fichero no encontrado: $scriptPath" -ForegroundColor Red; exit 1
        }
        $sqlContent = Get-Content $scriptPath -Raw -Encoding UTF8
        $statements = Split-SqlStatements -SqlBlock $sqlContent
        if (-not $Json) {
            Write-Host "  Fichero: $scriptPath" -ForegroundColor White
            Write-Host "  Sentencias: $($statements.Count)" -ForegroundColor White
            Write-Host ''
        }
    } else {
        $statements = Split-SqlStatements -SqlBlock $SqlInput
    }

    if ($statements.Count -eq 0) {
        Write-Host 'ERROR: No se encontraron sentencias SQL validas.' -ForegroundColor Red; exit 1
    }

    $tag = ''
    if ($Mode -in @('Seed','Teardown')) {
        $tag = if ($FixtureTag) { $FixtureTag } else { 'FX_' + (Get-Date -Format 'yyyyMMdd_HHmmss') }
    }

    $label = switch ($Mode) {
        'Exec'     { 'EXEC' }
        'Script'   { "SCRIPT ($([System.IO.Path]::GetFileName($Script)))" }
        'Seed'     { "SEED [$tag]" }
        'Teardown' { "TEARDOWN [$tag]" }
        'DDL'      { 'DDL' }
    }

    $report = Invoke-WriteStatements `
        -Statements $statements `
        -BackendInfo $t `
        -IsDryRun:$DryRun `
        -Label $label `
        -FixtureTagValue $tag `
        -OutputJson:$Json

    if ($Json) { $report | ConvertTo-Json -Depth 5 | Write-Output }

    if ($report.aborted) { exit 1 }
    exit 0
}

# =============================================================================
# HELP
# =============================================================================

Write-Host @"
ACCESS-QUERY v3 -- Consultas y escritura segura a backends Access (.accdb)

LECTURA:
  -SQL "SELECT ..."            SELECT libre (-Top 20 por defecto; -Top 0 = sin limite)
  -GetSchema -Table TbX        Esquema de campos (tipos, nullable)
  -Count -Table TbX            Contar registros
  -Distinct -Table Tb -Field C Valores unicos de un campo
  -ListTables                  Listar tablas locales
  -LinkedTables                Listar tablas linked (externas)
  -Compare -CompareSQL "..." -Backend A -CompareBackend B

ESCRITURA (con guardas):
  -Exec "SQL"                  SQL inline (multi-sentencia con ;)
  -Script "ruta.sql"           Desde fichero .sql

FIXTURES (requieren -AllowTable):
  -Seed   (-Exec "SQL" | -Script "seed.sql") -AllowTable "TbX" [-FixtureTag "TAG"]
  -Teardown (-Exec "SQL" | -Script "clean.sql") -AllowTable "TbX"

DDL:
  -CreateTable -Exec "CREATE TABLE ..."
  -DropTable -Table TbTest

SEGURIDAD:
  -DryRun            Validar sin ejecutar
  -AllowTable "X,Y"  Solo estas tablas aceptan escritura
  -DenyTable "A,B"   Bloquear adicionales (suma a backends.json)
  -StrictWrite       Requiere -AllowTable en TODO modo escritura
  -Force             Omitir -AllowTable en Seed/Teardown

SALIDA:
  -Json              JSON estructurado a stdout (funciona en lectura Y escritura)

PASSWORD (prioridad, las BDs sin password no necesitan configuracion):
  -Password > env ACCESS_QUERY_PW_<BACKEND> > env ACCESS_QUERY_PASSWORD > .secrets.json > backends.json

BACKENDS: $($backendMap.Keys -join ', ') | Default: $defaultBackend
"@ -ForegroundColor White
