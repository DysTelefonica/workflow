param(
    # -- Modos de lectura --
    [Parameter(Mandatory=$false)] [string]$SQL,
    [Parameter(Mandatory=$false)] [string]$Table,
    [Parameter(Mandatory=$false)] [string]$Field,
    [Parameter(Mandatory=$false)] [int]$Top,
    [Parameter(Mandatory=$false)] [switch]$Count,
    [Parameter(Mandatory=$false)] [switch]$Distinct,
    [Parameter(Mandatory=$false)] [switch]$ListTables,
    [Parameter(Mandatory=$false)] [switch]$LinkedTables,
    [Parameter(Mandatory=$false)] [switch]$GetSchema,
    [Parameter(Mandatory=$false)] [switch]$Compare,
    [Parameter(Mandatory=$false)] [string]$CompareBackend,
    [Parameter(Mandatory=$false)] [string]$CompareSQL,

    # -- Modos de escritura --
    [Parameter(Mandatory=$false)] [string]$Exec,
    [Parameter(Mandatory=$false)] [string]$Script,
    [Parameter(Mandatory=$false)] [switch]$Seed,
    [Parameter(Mandatory=$false)] [switch]$Teardown,
    [Parameter(Mandatory=$false)] [string]$FixtureTag,
    [Parameter(Mandatory=$false)] [switch]$CreateTable,
    [Parameter(Mandatory=$false)] [switch]$DropTable,

    # -- Guardas de seguridad --
    [Parameter(Mandatory=$false)] [switch]$DryRun,
    [Parameter(Mandatory=$false)] [string]$AllowTable,
    [Parameter(Mandatory=$false)] [string]$DenyTable,
    [Parameter(Mandatory=$false)] [switch]$StrictWrite,
    [Parameter(Mandatory=$false)] [switch]$Force,

    # -- Salida --
    [Parameter(Mandatory=$false)] [switch]$Json,

    # -- Conexion --
    [Parameter(Mandatory=$false)] [string]$Backend = '',
    [Parameter(Mandatory=$false)] [string]$BackendPath = '',
    [Parameter(Mandatory=$false)] [string]$Password = ''
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# FIX #0: VALIDACION EXPLÍCITA DE COMBINACIONES INCOMPATIBLES
# Fail-fast con mensajes claros, antes del dispatcher.
# =============================================================================
$validationErrors = @()

# -- Solo un modo PRINCIPAL (selector de modo) a la vez --
# Los switches de modo (-Seed, -Teardown) + parámetros con contenido (-SQL, -Exec, etc)
$sqlModeProvided      = $SQL -ne ''
$execModeProvided     = $Exec -ne ''
$scriptModeProvided   = $Script -ne ''
$compareModeProvided  = $CompareSQL -ne ''  # Compare se activa con -CompareSQL, no solo -Compare
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

$writeModes = @($execModeProvided, $scriptModeProvided, $Seed, $Teardown, $createTableProvided, $dropTableProvided)
$activeWriteModes = ($writeModes | Where-Object { $_ }).Count

# Solo un read mode
if ($activeReadModes -gt 1) {
    $validationErrors += "Solo puedes usar un modo de lectura a la vez (-GetSchema, -Count, -Distinct, -Compare, -ListTables, -LinkedTables)."
}

# Solo un write mode (que no sea Seed/Teardown que tienen su propia guarda)
if ($execModeProvided -and $scriptModeProvided) {
    $validationErrors += "-Exec y -Script son mutuamente excluyentes (elige uno)."
}

# -- -SQL incompatible con modos de escritura Y lectura avanzada --
if ($sqlModeProvided -and ($execModeProvided -or $scriptModeProvided)) {
    $validationErrors += "-SQL es modo de solo lectura: no puede combinarse con -Exec o -Script."
}
if ($sqlModeProvided -and $compareModeProvided) {
    $validationErrors += "-SQL y -Compare son mutuamente excluyentes (son ambos de solo lectura, pero incompatibles)."
}
if ($sqlModeProvided -and $getSchemaProvided) {
    $validationErrors += "-SQL y -GetSchema son mutuamente excluyentes."
}
if ($sqlModeProvided -and $countProvided) {
    $validationErrors += "-SQL y -Count son mutuamente excluyentes."
}
if ($sqlModeProvided -and $distinctProvided) {
    $validationErrors += "-SQL y -Distinct son mutuamente excluyentes."
}
if ($sqlModeProvided -and $listTablesProvided) {
    $validationErrors += "-SQL y -ListTables son mutuamente excluyentes."
}
if ($sqlModeProvided -and $linkedTablesProvided) {
    $validationErrors += "-SQL y -LinkedTables son mutuamente excluyentes."
}

# -- -GetSchema no puede usar -Exec --
if ($getSchemaProvided -and $execModeProvided) {
    $validationErrors += "-GetSchema es modo de solo lectura: no puede combinarse con -Exec."
}

# -- -Compare (con -CompareSQL) no puede mezclarse con modos de escritura --
if ($Compare -and ($execModeProvided -or $scriptModeProvided -or $Seed -or $Teardown)) {
    $validationErrors += "-Compare es modo de solo lectura: no puede combinarse con -Exec, -Script, -Seed o -Teardown."
}

# -- Seed/Teardown incompatibilidades --
if ($Seed -and $Teardown) {
    $validationErrors += "-Seed y -Teardown son mutuamente excluyentes. Elegir uno."
}
if ($Seed -and $createTableProvided) {
    $validationErrors += "-Seed no puede combinarse con -CreateTable (son DDL, no fixtures)."
}
if ($Seed -and $dropTableProvided) {
    $validationErrors += "-Seed no puede combinarse con -DropTable (son DDL, no fixtures)."
}
if ($Teardown -and $createTableProvided) {
    $validationErrors += "-Teardown no puede combinarse con -CreateTable (son DDL, no fixtures)."
}
if ($Teardown -and $dropTableProvided) {
    $validationErrors += "-Teardown no puede combinarse con -DropTable (son DDL, no fixtures)."
}

# -- DDL incompatibility --
if ($createTableProvided -and $dropTableProvided) {
    $validationErrors += "-DropTable y -CreateTable son mutuamente excluyentes."
}

# -- Mostrar errores y abortar --
if ($validationErrors.Count -gt 0) {
    foreach ($err in $validationErrors) {
        Write-Host "ERROR: $err" -ForegroundColor Red
    }
    Write-Host "Ejecuta '.\query-backend.ps1' sin argumentos para ver la ayuda." -ForegroundColor Yellow
    exit 1
}

# =============================================================================
# FIX #1: DISPATCHER CENTRALIZADO
# Resolver el modo UNA SOLA VEZ. Los modos compuestos (Seed+Exec, Teardown+Script)
# se evaluan ANTES que los simples para evitar que -Exec absorba -Seed -Exec.
# =============================================================================

$Mode = $null
$SqlInput = $null

# -- Modos compuestos (mas especificos primero) --
if     ($Seed -and $Exec)            { $Mode = 'Seed';      $SqlInput = $Exec }
elseif ($Seed -and $Script)          { $Mode = 'Seed';      $SqlInput = '__FILE__' }
elseif ($Teardown -and $Exec)        { $Mode = 'Teardown';  $SqlInput = $Exec }
elseif ($Teardown -and $Script)      { $Mode = 'Teardown';  $SqlInput = '__FILE__' }
elseif ($CreateTable -and $Exec)     { $Mode = 'DDL';       $SqlInput = $Exec }
elseif ($DropTable -and $Table)      { $Mode = 'DDL';       $SqlInput = "DROP TABLE [$Table]" }
# -- Modos simples de escritura --
elseif ($Exec)                       { $Mode = 'Exec';      $SqlInput = $Exec }
elseif ($Script)                     { $Mode = 'Script';    $SqlInput = '__FILE__' }
# -- Modos de lectura --
elseif ($ListTables)                 { $Mode = 'ListTables' }
elseif ($LinkedTables)               { $Mode = 'LinkedTables' }
elseif ($GetSchema -and $Table)      { $Mode = 'GetSchema' }
elseif ($Count -and $Table)          { $Mode = 'Count' }
elseif ($Distinct -and $Table -and $Field) { $Mode = 'Distinct' }
elseif ($Compare -and $CompareSQL)   { $Mode = 'Compare' }
elseif ($SQL)                        { $Mode = 'SQL' }
# -- Validaciones de combinaciones invalidas --
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

# -- Deny-list global: backends.json + parametro CLI --
$globalDenyTables = @()
if ($config.PSObject.Properties['deny_tables']) {
    $globalDenyTables = @($config.deny_tables)
}
if ($DenyTable) {
    $globalDenyTables += ($DenyTable -split ',') | ForEach-Object { $_.Trim() }
}
$globalDenyTables = $globalDenyTables | Select-Object -Unique

# -- Allow-list desde parametro --
$allowTableList = @()
if ($AllowTable) {
    $allowTableList = ($AllowTable -split ',') | ForEach-Object { $_.Trim() }
}

# =============================================================================
# FIX #2: RESOLUCION DE PASSWORDS SIN HARDCODING
# Cadena de prioridad:
#   1. -Password (CLI override)
#   2. Env var ACCESS_QUERY_PW_<BACKEND>
#   3. Env var ACCESS_QUERY_PASSWORD (global)
#   4. .secrets.json (fichero local, no versionar)
#   5. backends.json > password (backward compat, DEPRECADO)
#   6. Error claro
# =============================================================================

function Resolve-Password {
    param(
        [string]$BackendName,
        [string]$CliPassword,
        [string]$JsonPassword
    )
    # 1. CLI override
    if ($CliPassword) { return $CliPassword }

    # 2. Env var por backend
    $envPerBackend = "ACCESS_QUERY_PW_$($BackendName -replace '[^a-zA-Z0-9]','_')"
    $envVal = [System.Environment]::GetEnvironmentVariable($envPerBackend)
    if ($envVal) { return $envVal }

    # 3. Env var global
    $envGlobal = [System.Environment]::GetEnvironmentVariable('ACCESS_QUERY_PASSWORD')
    if ($envGlobal) { return $envGlobal }

    # 4. .secrets.json
    $secretsPath = Join-Path $ScriptDir '.secrets.json'
    if (Test-Path $secretsPath) {
        try {
            $secrets = Get-Content $secretsPath -Raw | ConvertFrom-Json
            if ($secrets.PSObject.Properties[$BackendName]) {
                return $secrets.$BackendName
            }
            if ($secrets.PSObject.Properties['default']) {
                return $secrets.default
            }
        } catch { }
    }

    # 5. backends.json (backward compat)
    if ($JsonPassword) { return $JsonPassword }

    # 6. Sin password
    return ''
}

# =============================================================================
# FUNCIONES UTILITARIAS
# =============================================================================

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
        $resolvedPw = Resolve-Password -BackendName '__direct__' -CliPassword $Password -JsonPassword ''
        return @{ Path = $OverridePath; Password = $resolvedPw; Name = (Split-Path $OverridePath -Leaf) }
    }
    if ($Name -eq '') { $Name = $defaultBackend }
    if (-not $backendMap.ContainsKey($Name)) {
        Write-Host "ERROR: Backend '$Name' no encontrado. Disponibles: $($backendMap.Keys -join ', ')" -ForegroundColor Red; exit 1
    }
    $info = $backendMap[$Name]
    $jsonPw = if ($info.PSObject.Properties['password']) { $info.password } else { '' }
    $resolvedPw = Resolve-Password -BackendName $Name -CliPassword $Password -JsonPassword $jsonPw
    if (-not $resolvedPw) {
        Write-Host "ERROR: No se encontro password para backend '$Name'." -ForegroundColor Red
        Write-Host '  Opciones: -Password "pw", env ACCESS_QUERY_PW_<BACKEND>, .secrets.json, o backends.json' -ForegroundColor Yellow
        exit 1
    }
    return @{ Path = $info.path; Password = $resolvedPw; Name = $Name }
}

function Format-Value {
    param($val)
    if ($val -is [System.DBNull] -or $null -eq $val) { return 'NULL' }
    if ($val -is [string] -and $val.Length -gt 50) { return $val.Substring(0, 47) + '...' }
    return [string]$val
}

# =============================================================================
# FIX #3: PARSER DE SENTENCIAS ROBUSTO
# State machine que respeta ; dentro de strings (' y ") y comentarios --.
# =============================================================================

function Split-SqlStatements {
    param([string]$SqlBlock)
    $statements = [System.Collections.Generic.List[string]]::new()
    $current = [System.Text.StringBuilder]::new()
    $inSingleQuote = $false
    $inDoubleQuote = $false

    for ($i = 0; $i -lt $SqlBlock.Length; $i++) {
        $c = $SqlBlock[$i]

        # -- Comilla simple: toggle, pero '' es escape en Access SQL --
        if ($c -eq "'" -and -not $inDoubleQuote) {
            if ($inSingleQuote -and ($i + 1) -lt $SqlBlock.Length -and $SqlBlock[$i + 1] -eq "'") {
                [void]$current.Append($c)
                $i++
                [void]$current.Append($SqlBlock[$i])
                continue
            }
            $inSingleQuote = -not $inSingleQuote
            [void]$current.Append($c)
            continue
        }

        # -- Comilla doble: toggle --
        if ($c -eq '"' -and -not $inSingleQuote) {
            $inDoubleQuote = -not $inDoubleQuote
            [void]$current.Append($c)
            continue
        }

        # -- Comentario inline -- (solo fuera de strings) --
        if ($c -eq '-' -and -not $inSingleQuote -and -not $inDoubleQuote) {
            if (($i + 1) -lt $SqlBlock.Length -and $SqlBlock[$i + 1] -eq '-') {
                # Saltar hasta fin de linea
                while ($i -lt $SqlBlock.Length -and $SqlBlock[$i] -ne "`n") { $i++ }
                continue
            }
        }

        # -- Separador ; solo fuera de strings --
        if ($c -eq ';' -and -not $inSingleQuote -and -not $inDoubleQuote) {
            $stmt = $current.ToString().Trim()
            if ($stmt) { [void]$statements.Add($stmt) }
            [void]$current.Clear()
            continue
        }

        [void]$current.Append($c)
    }

    # Ultima sentencia sin ; final
    $lastStmt = $current.ToString().Trim()
    if ($lastStmt) { [void]$statements.Add($lastStmt) }

    return $statements.ToArray()
}

# =============================================================================
# FUNCIONES DE SEGURIDAD
# =============================================================================

function Get-LinkedTableNames {
    param([System.Data.OleDb.OleDbConnection]$Conn)
    $linked = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $schema = $Conn.GetSchema('Tables')
    foreach ($row in $schema) {
        if ($row.Item('TABLE_TYPE') -eq 'LINK') {
            [void]$linked.Add($row.Item('TABLE_NAME'))
        }
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
            if ($tbl -like $deny) {
                $blocked += "DENY: '$tbl' coincide con pattern '$deny'"
            }
        }
        if ($LinkedTables.Contains($tbl)) {
            $blocked += "LINKED: '$tbl' es tabla LINKED/EXTERNA"
        }
        if ($AllowList.Count -gt 0) {
            $allowed = $false
            foreach ($allow in $AllowList) {
                if ($tbl -like $allow) { $allowed = $true; break }
            }
            if (-not $allowed) {
                $blocked += "ALLOW: '$tbl' no esta en allow-list ($($AllowList -join ', '))"
            }
        }
    }
    return @{ Targets = $targets; Blocked = $blocked }
}

# =============================================================================
# FIX #6: VALIDACION ESTRICTA PARA SEED/TEARDOWN
# -Seed y -Teardown REQUIEREN -AllowTable (a menos que -Force).
# -StrictWrite extiende esto a todos los modos de escritura.
# =============================================================================

$isWriteMode = $Mode -in @('Exec','Script','Seed','Teardown','DDL')

if ($isWriteMode) {
    if ($Mode -in @('Seed','Teardown') -and $allowTableList.Count -eq 0 -and -not $Force) {
        Write-Host "ERROR: -$Mode requiere -AllowTable para evitar tocar tablas equivocadas." -ForegroundColor Red
        Write-Host '  Ejemplo: -AllowTable "TbSolicitudes,TbDocumentos"' -ForegroundColor Yellow
        Write-Host '  Usa -Force para saltarte esta restriccion (no recomendado).' -ForegroundColor DarkYellow
        exit 1
    }
    if ($StrictWrite -and $allowTableList.Count -eq 0 -and -not $Force) {
        Write-Host 'ERROR: -StrictWrite activo: se requiere -AllowTable explicito.' -ForegroundColor Red
        Write-Host '  Ejemplo: -AllowTable "TbSolicitudes"' -ForegroundColor Yellow
        exit 1
    }
}

# =============================================================================
# FIX #4 + #5: MOTOR DE EJECUCION CON FIXTURE TRACKING + SALIDA JSON
# - Fixture log real: .fixture-log.json acumulativo
# - -Json emite JSON estructurado a stdout (Write-Output)
# - Write-Host para humanos (no interfiere con JSON)
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

    # Estructura de resultado
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
                    Write-Host ''
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
                            $cols += "$($reader.GetName($ci))=$(Format-Value $reader.GetValue($ci))"
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

    # -- Resumen humano --
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
        if ($allTablesWritten.Count -gt 0) {
            Write-Host "  Tablas: $($allTablesWritten -join ', ')" -ForegroundColor Yellow
        }
        if (-not $IsDryRun) {
            Write-Host "  Filas afectadas: $($report.totalAffected)" -ForegroundColor Green
        }
    }

    # -- Fixture log: acumular en .fixture-log.json --
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
        if (-not $OutputJson) {
            Write-Host "  Fixture log: $logPath" -ForegroundColor DarkGray
        }
    }

    return $report
}

# =============================================================================
# EJECUCION POR MODO
# =============================================================================

# -- Modos de lectura (funcionalmente identicos al original) --

if ($Mode -eq 'ListTables') {
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

if ($Mode -eq 'LinkedTables') {
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

if ($Mode -eq 'GetSchema') {
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
    Write-Host "  | $("Campo".PadRight($colWidth)) | Tipo           | Nullable |" -ForegroundColor White
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

if ($Mode -eq 'Count') {
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
    Write-Host "=== DISTINCT $Field ON $Table ($($t.Name)) ===" -ForegroundColor Cyan
    Write-Host "  $($vals.Count) valores:" -ForegroundColor Gray
    $vals | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    exit 0
}

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

if ($Mode -eq 'SQL') {
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

# -- Modos de escritura (unificados) --

if ($Mode -in @('Exec','Script','Seed','Teardown','DDL')) {
    $t = Get-BackendPath -Name $Backend -OverridePath $BackendPath

    # Resolver SQL input
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

    # Resolver fixture tag
    $tag = ''
    if ($Mode -in @('Seed','Teardown')) {
        $tag = if ($FixtureTag) { $FixtureTag } else { 'FX_' + (Get-Date -Format 'yyyyMMdd_HHmmss') }
    }

    # Resolver label
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

    # FIX #5: Salida JSON estructurada a stdout
    if ($Json) {
        $report | ConvertTo-Json -Depth 5 | Write-Output
    }

    if ($report.aborted) { exit 1 }
    exit 0
}

# =============================================================================
# HELP
# =============================================================================

Write-Host @"
ACCESS-QUERY v2 -- Consultas y escritura segura a backends Access (.accdb)

LECTURA:
  -SQL "SELECT ..."            SELECT libre (-Top N para limitar)
  -GetSchema -Table TbX        Esquema de campos
  -Count -Table TbX            Contar registros
  -Distinct -Table Tb -Field C Valores unicos
  -ListTables / -LinkedTables  Listar tablas
  -Compare -CompareSQL "..." -Backend A -CompareBackend B

ESCRITURA (con guardas):
  -Exec "SQL"                  SQL inline (multi-sentencia con ;)
  -Script "ruta.sql"           Desde fichero

FIXTURES (requieren -AllowTable):
  -Seed -Exec "SQL" -AllowTable "TbX" [-FixtureTag "TAG"]
  -Seed -Script "seed.sql" -AllowTable "TbX"
  -Teardown -Exec "SQL" -AllowTable "TbX"
  -Teardown -Script "clean.sql" -AllowTable "TbX"

DDL:
  -CreateTable -Exec "CREATE TABLE ..."
  -DropTable -Table TbTest

SEGURIDAD:
  -DryRun            Validar sin ejecutar
  -AllowTable "X,Y"  Solo estas tablas
  -DenyTable "A,B"   Bloquear (suma a backends.json)
  -StrictWrite       Requiere AllowTable en TODO modo escritura
  -Force             Omitir AllowTable en Seed/Teardown

SALIDA:
  -Json              JSON estructurado a stdout

PASSWORD (prioridad):
  -Password > env ACCESS_QUERY_PW_<BACKEND> > env ACCESS_QUERY_PASSWORD > .secrets.json > backends.json

BACKENDS: $($backendMap.Keys -join ', ') | Default: $defaultBackend
"@ -ForegroundColor White
