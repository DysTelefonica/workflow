[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del skill.")]
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("New-Sandbox", "Discover-LinkedTables", "Make-Sidecar", "Localize-Sandbox")]
    [string]$Action,

    [Parameter()]
    [string]$SourceBackend,

    [Parameter()]
    [string]$SandboxPath,

    [Parameter()]
    [string]$Password = "",

    [Parameter()]
    [string]$SidecarSuffix = "_nopass",

    [Parameter()]
    [string]$SourceSidecar,

    [Parameter()]
    [string]$SandboxPassword = "",
    
    # BUG 4 FIX: Hashtable for multi-backend scenario — passed as dictionary string
    # Example: @{ "\\server\share\Expedientes.accdb"="C:\sidecar\Expedientes_nopass.accdb" }
    # On command line use: -BackendSidecarMap (@{ "path1"="sidecar1" } | ConvertTo-Json -Compress)
    [Parameter()]
    [string]$BackendSidecarMapJson = ""
)

$ErrorActionPreference = "Stop"

# ============================================================
# HELPERS — COMUNES
# ============================================================

function Write-Status {
    Param(
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

# Verifica que DAO.DBEngine esté disponible (DLL in-process, sin MSACCESS.EXE)
# Esto es más preciso que Test-AccessInstalled porque DAO.DBEngine viene con
# el Access Database Engine Redistributable (sin Access completo instalado)
function Test-DaoAvailable {
    try {
        $engine = New-DaoDbEngine
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        return $true
    } catch {
        return $false
    }
}

function New-DaoDbEngine {
    $engineCandidates = @(
        "DAO.DBEngine.160",
        "DAO.DBEngine.150",
        "DAO.DBEngine.140",
        "DAO.DBEngine.120",
        "DAO.DBEngine.36"
    )

    foreach ($progId in $engineCandidates) {
        try {
            $engine = New-Object -ComObject $progId
            if ($engine) { return $engine }
        } catch {}
    }

    throw "No se pudo crear ninguna versión de DAO.DBEngine (120/140/150/160)."
}

function Connect-Database {
    Param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Password = "",
        [switch]$Exclusive = $false
    )

    $engine = New-DaoDbEngine
    $db = $null

    try {
        $connect = ""
        if (-not [string]::IsNullOrEmpty($Password)) {
            $connect = ";PWD=$Password"
        }

        $db = $engine.OpenDatabase($Path, $Exclusive, $false, $connect)
        return [pscustomobject]@{
            Engine = $engine
            Database = $db
        }
    } catch {
        if ($db) { try { $db.Close() } catch {} }
        if ($engine) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null } catch {} }
        throw
    }
}

function Close-DaoObjects {
    Param([Parameter(Mandatory = $true)]$Objects)
    foreach ($obj in $Objects) {
        if ($obj) {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {}
        }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Get-ProcessIdFromHwnd {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][IntPtr]$Hwnd
    )

    if (-not ([System.Management.Automation.PSTypeName]"Win32.NativeMethods").Type) {
        Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @"
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
    }

    [uint32]$pid = 0
    [Win32.NativeMethods]::GetWindowThreadProcessId($Hwnd, [ref]$pid) | Out-Null
    return [int]$pid
}

function Open-AccessSession {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password = ""
    )

    $access = $null
    $accessPid = $null
    $prePids = @()

    try {
        # Registrar PIDs de Access existentes antes de abrir (para identificar el nuestro)
        try {
            $prePids = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        } catch {
            $prePids = @()
        }

        # Crear instancia COM de Access — modo unattended
        $access = New-Object -ComObject Access.Application
        $access.Visible = $false
        $access.UserControl = $false
        $access.AutomationSecurity = 1

        # Obtener PID via hWnd
        try {
            $hwnd = [IntPtr]$access.hWndAccessApp
            if ($hwnd -and $hwnd -ne [IntPtr]::Zero) {
                $accessPid = Get-ProcessIdFromHwnd -Hwnd $hwnd
            }
        } catch {}

        # Abrir la base de datos
        $access.OpenCurrentDatabase($AccessPath, $false, $Password)

        # Suprimir TODOS los warnings/prompts de Access — CRÍTICO para modo unattended
        try { $access.DoCmd.SetWarnings($false) } catch {}

        # Si no se obtuvo PID por hWnd, buscar el proceso nuevo
        if (-not $accessPid) {
            try {
                $post = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -Property Id, StartTime)
                $new = @($post | Where-Object { $_.Id -notin $prePids })
                if ($new.Count -ge 1) {
                    $picked = $new | Sort-Object -Property StartTime -Descending | Select-Object -First 1
                    $accessPid = [int]$picked.Id
                }
            } catch {}
        }

        return [pscustomobject]@{
            AccessApplication = $access
            ProcessId         = $accessPid
            PrePids           = $prePids
        }

    } catch {
        # Cleanup en caso de error durante la apertura
        if ($access) {
            try { $access.CloseCurrentDatabase() } catch {}
            try { $access.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) | Out-Null } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        throw
    }
}

function Close-AccessSession {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Session
    )

    $access = $Session.AccessApplication
    $accessPid = $Session.ProcessId

    if ($access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
    }

    if ($access) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) | Out-Null } catch {}
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Solo matar el proceso si es nuestro (el que abrimos en esta sesión)
    if ($accessPid) {
        try {
            $stillAlive = Get-Process -Id $accessPid -ErrorAction SilentlyContinue
            if ($stillAlive) {
                Stop-Process -Id $accessPid -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}

function Get-TableDefs {
    Param(
        [Parameter(Mandatory = $true)]$Database
    )

    $result = @()
    $tableDefs = $Database.TableDefs

    for ($i = 0; $i -lt $tableDefs.Count; $i++) {
        $td = $tableDefs[$i]
        try {
            $name = $td.Name
            if ($name -match "^MSys" -or $name -match "^~") { continue }

            $connect = ""
            try { $connect = $td.Connect } catch {}

            $isLinked = $false
            $sourceTable = $name
            if (-not [string]::IsNullOrEmpty($connect) -and $connect -match ";DATABASE=") {
                $isLinked = $true
                if ($connect -match ";DATABASE=([^;]+)") {
                    $sourceTable = $name  # linked table name is the same
                }
            }

            $result += [pscustomobject]@{
                TableName    = $name
                Connect      = $connect
                SourceTable  = $sourceTable
                IsLinked     = $isLinked
                SourcePath   = if ($isLinked -and $connect -match ";DATABASE=([^;]+)") { $Matches[1].Trim() } else { $null }
            }
        } catch {}
    }

    return $result
}

function Copy-TableData {
    Param(
        [Parameter(Mandatory = $true)]$SourcePath,
        [Parameter(Mandatory = $true)]$DestDb,
        [Parameter(Mandatory = $true)][string]$TableName,
        [string]$Password = ""
    )

    # Use a separate connection for the source to allow it to stay open
    $sourceConn = $null
    $sourceEngine = $null
    $sourceDbObj = $null

    try {
        $sourceEngine = New-DaoDbEngine
        $sourceConnect = if ($Password) { ";PWD=$Password" } else { "" }
        $sourceDbObj = $sourceEngine.OpenDatabase($SourcePath, $false, $false, $sourceConnect)
        $sourceConn = $sourceDbObj.TableDefs[$TableName]

        if (-not $sourceConn) {
            Write-Status -Message "WARNING: Tabla origen '$TableName' no encontrada" -Color Yellow
            return
        }

        # Get field definitions from source (BUG N FIX: preserve Attributes for AutoNumber/Required)
        $fields = @()
        for ($j = 0; $j -lt $sourceConn.Fields.Count; $j++) {
            $f = $sourceConn.Fields[$j]
            $fields += [pscustomobject]@{
                Name = $f.Name
                Type = $f.Type
                Size = $f.Size
                Attributes = $f.Attributes  # BUG N FIX: preserve AutoNumber, Required, etc.
            }
        }

        # Create destination table
        $destTd = $destDb.CreateTableDef($TableName)

        foreach ($fieldDef in $fields) {
            try {
                $newField = $destTd.CreateField($fieldDef.Name, $fieldDef.Type, $fieldDef.Size)
                $newField.Attributes = $fieldDef.Attributes  # BUG N FIX: restore Attributes
                $destTd.Fields.Append($newField)
            } catch {
                Write-Status -Message "WARNING: No se pudo crear campo '$($fieldDef.Name)': $($_.Exception.Message)" -Color Yellow
            }
        }

        $destDb.TableDefs.Append($destTd)

        # Copy data using recordset
        $sourceRs = $sourceDbObj.OpenRecordset($TableName, 1)  # dbOpenTable = 1
        $destRs = $destDb.OpenRecordset($TableName, 2)  # dbOpenDynaset = 2

        try {
            while (-not $sourceRs.EOF) {
                $destRs.AddNew
                for ($k = 0; $k -lt $sourceRs.Fields.Count; $k++) {
                    try {
                        $destRs.Fields[$k].Value = $sourceRs.Fields[$k].Value
                    } catch {}
                }
                $destRs.Update
                $sourceRs.MoveNext
            }
        } finally {
            try { $sourceRs.Close() } catch {}
            try { $destRs.Close() } catch {}
        }

    } finally {
        if ($sourceDbObj) { try { $sourceDbObj.Close() } catch {} }
        Close-DaoObjects -Objects @($sourceDbObj, $sourceEngine)
    }
}

function New-Sandbox {
    Param(
        [Parameter(Mandatory = $true)][string]$SourceBackend,
        [Parameter(Mandatory = $true)][string]$SandboxPath
    )

    if (-not (Test-Path -Path $SourceBackend)) {
        throw "SourceBackend no encontrado: $SourceBackend"
    }

    # Resolve and validate sandbox path
    $sandboxRoot = Split-Path -Parent $SandboxPath
    if (-not (Test-Path -Path $sandboxRoot)) {
        New-Item -Path $sandboxRoot -ItemType Directory -Force | Out-Null
    }

    Write-Status -Message "Copiando backend a sandbox..." -Color Cyan
    Write-Status -Message "  Fuente:     $SourceBackend"
    Write-Status -Message "  Destino:    $SandboxPath"

    Copy-Item -Path $SourceBackend -Destination $SandboxPath -Force

    $srcSize = (Get-Item $SourceBackend).Length
    $dstSize = (Get-Item $SandboxPath).Length

    if ($dstSize -eq 0) {
        throw "CRITICAL: El archivo copiado tiene tamaño 0"
    }

    Write-Status -Message "OK Sandbox creado ($([math]::Round($dstSize/1KB, 1)) KB)" -Color Green
}

function Discover-LinkedTables {
    Param(
        [Parameter(Mandatory = $true)][string]$SandboxPath,
        [string]$Password = ""
    )

    if (-not (Test-Path -Path $SandboxPath)) {
        throw "SandboxPath no encontrado: $SandboxPath"
    }

    Write-Status -Message "Descubriendo tablas vinculadas en sandbox..." -Color Cyan

    $conn = $null
    $engine = $null
    $db = $null

    try {
        $engine = New-DaoDbEngine
        $connect = if ($Password) { ";PWD=$Password" } else { "" }
        $db = $engine.OpenDatabase($SandboxPath, $false, $false, $connect)

        $tables = Get-TableDefs -Database $db
        $linked = $tables | Where-Object { $_.IsLinked }

        if ($linked.Count -eq 0) {
            Write-Status -Message "No se encontraron tablas vinculadas" -Color Green
        } else {
            Write-Status -Message "Encontradas $($linked.Count) tabla(s) vinculada(s):" -Color Cyan
            foreach ($t in $linked) {
                Write-Status -Message "  - $($t.TableName) -> $($t.SourcePath)" -Color Gray
            }
        }

        return $linked

    } finally {
        if ($db) { try { $db.Close() } catch {} }
        Close-DaoObjects -Objects @($db, $engine)
    }
}

function Make-Sidecar {
    Param(
        [Parameter(Mandatory = $true)][string]$SourceBackend,
        [string]$Password = "",
        [string]$SidecarSuffix = "_nopass"
    )

    if (-not (Test-Path -Path $SourceBackend)) {
        throw "SourceBackend no encontrado: $SourceBackend"
    }

    $extension = [System.IO.Path]::GetExtension($SourceBackend)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourceBackend)
    $baseDir = Split-Path -Parent $SourceBackend
    $sidecarPath = Join-Path -Path $baseDir -ChildPath ($baseName + $SidecarSuffix + $extension)

    Write-Status -Message "Creando sidecar sin password..." -Color Cyan
    Write-Status -Message "  Fuente:    $SourceBackend"
    Write-Status -Message "  Sidecar:   $sidecarPath"

    # Open source with password
    $srcEngine = $null
    $srcDb = $null

    try {
        $srcEngine = New-DaoDbEngine
        $srcConnect = if ($Password) { ";PWD=$Password" } else { "" }
        $srcDb = $srcEngine.OpenDatabase($SourceBackend, $false, $false, $srcConnect)

        # Get all table definitions (local and linked)
        $allTables = Get-TableDefs -Database $srcDb

        if ($allTables.Count -eq 0) {
            Write-Status -Message "WARNING: No se encontraron tablas en el backend" -Color Yellow
            return $sidecarPath
        }

        # Create destination as a fresh ACCDB
        # We use DAO to create a new database
        $tmpMdb = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("sidecar_tmp_{0}.accdb" -f [guid]::NewGuid().ToString("N"))

        # DAO can create via DBEngine.CreateDatabase
        # Use dbLangGeneral = "" (general text, LCID 0x0409) and dbVersion120 = 128 (Access 2007+ .accdb)
        $destEngine = New-DaoDbEngine
        try {
            $destLocale = ""  # dbLangGeneral = ""
            $dbVersion = 128  # dbVersion120 = 0x80 = 128 (Access 2007+ .accdb format)
            $destDbNew = $destEngine.CreateDatabase($tmpMdb, $destLocale, $dbVersion)
            $destDbNew.Close()
        } finally {
            Close-DaoObjects -Objects @($destDbNew, $destEngine)
        }

        # Now open the newly created DB and copy all tables
        $destConn = $null
        $destEngine2 = $null
        $destDb2 = $null

        try {
            $destEngine2 = New-DaoDbEngine
            $destDb2 = $destEngine2.OpenDatabase($tmpMdb, $false, $false, "")

            foreach ($tbl in $allTables) {
                # BUG G FIX: Use $tbl.TableName instead of $tbl.Name (which is always $null)
                if ($tbl.TableName -match "^MSys" -or $tbl.TableName -match "^~") { continue }

                Write-Status -Message "  Copiando tabla: $($tbl.TableName)..." -Color Gray

                if ($tbl.IsLinked) {
                    # BUG B FIX: Linked tables are skipped in sidecar creation
                    # The sidecar only contains local tables from the source backend.
                    # External linked backends are NOT opened (each may have different passwords).
                    Write-Status -Message "    SKIP: Tabla vinculada '$($tbl.TableName)' no se incluye en sidecar (requeriría acceso a backend externo)" -Color Yellow
                    continue
                } else {
                    # Local table: copy directly from source
                    Copy-TableData -SourcePath $SourceBackend -DestDb $destDb2 -TableName $tbl.TableName -Password $Password
                }
            }

        } finally {
            # .Close() antes de FinalReleaseComObject — necesario para flush de writes y liberar file lock
            if ($destDb2) { try { $destDb2.Close() } catch {} }
            Close-DaoObjects -Objects @($destDb2, $destEngine2)
        }

        # Move temp file to final destination
        Move-Item -Path $tmpMdb -Destination $sidecarPath -Force

        $size = (Get-Item $sidecarPath).Length
        Write-Status -Message "OK Sidecar creado ($([math]::Round($size/1KB, 1)) KB): $sidecarPath" -Color Green

        return $sidecarPath

    } finally {
        # .Close() antes de FinalReleaseComObject — necesario para flush de writes y liberar file lock
        if ($srcDb) { try { $srcDb.Close() } catch {} }
        Close-DaoObjects -Objects @($srcDb, $srcEngine)
    }
}

function Localize-Sandbox {
    Param(
        [Parameter(Mandatory = $true)][string]$SandboxPath,
        [string]$SourceSidecar = "",
        [string]$SourceBackend = "",
        [string]$Password = "",
        [string]$SandboxPassword = "",
        
        # BUG 4 FIX: Hashtable for multi-backend scenario
        # Maps source backend path -> sidecar path (e.g., for projects with multiple external backends)
        [hashtable]$BackendSidecarMap = $null
    )

    if (-not (Test-Path -Path $SandboxPath)) {
        throw "SandboxPath no encontrado: $SandboxPath"
    }

    # Determine which source to use for data
    $dataSource = $null
    $dataPassword = $null

    if ($SourceSidecar) {
        if (-not (Test-Path -Path $SourceSidecar)) {
            throw "SourceSidecar no encontrado: $SourceSidecar"
        }
        $dataSource = $SourceSidecar
        $dataPassword = ""  # sidecar has no password
    } elseif ($SourceBackend) {
        $dataSource = $SourceBackend
        $dataPassword = $Password
    } elseif ($BackendSidecarMap) {
        # BUG 4 FIX: Multi-backend mode — BackendSidecarMap takes precedence
        # The map will be used during table processing to resolve each linked table's source
        Write-Status -Message "Localizando sandbox en modo multi-backend ($($BackendSidecarMap.Count) backend(s))..." -Color Cyan
    } else {
        throw "Debe especificar -SourceSidecar, -SourceBackend, o -BackendSidecarMap"
    }

    Write-Status -Message "Localizando sandbox (reemplazando vínculos con tablas locales)..." -Color Cyan
    Write-Status -Message "  Sandbox:    $SandboxPath"
    if ($dataSource) {
        Write-Status -Message "  DataSource: $dataSource"
    }

    # Discover current linked tables
    $linkedTables = Discover-LinkedTables -SandboxPath $SandboxPath -Password $SandboxPassword

    if ($linkedTables.Count -eq 0) {
        Write-Status -Message "Sandbox ya está localizeada (sin vínculos). Nada que hacer." -Color Green
        return
    }

    # BUG 5 FIX: Removed Access.Application usage — keep purely DAO-based
    # The Access COM session was unnecessary overhead; all operations use DAO
    $db = $null
    $daoEngine = $null

    try {
        $daoEngine = New-DaoDbEngine
        $sandboxConnect = if ($SandboxPassword) { ";PWD=$SandboxPassword" } else { "" }
        $db = $daoEngine.OpenDatabase($SandboxPath, $false, $false, $sandboxConnect)

        foreach ($linked in $linkedTables) {
            Write-Status -Message "  Procesando tabla vinculada: $($linked.TableName)" -Color Gray

            $sourcePath = $linked.SourcePath
            $sourceTableName = $linked.SourceTable

            if (-not (Test-Path -Path $sourcePath)) {
                Write-Status -Message "    WARNING: Source no encontrado: $sourcePath. Saltando." -Color Yellow
                continue
            }

            # BUG 4 FIX: Resolve which sidecar/backend to use for this linked table
            # In multi-backend mode, each source backend has its own sidecar
            $tableDataSource = $null
            $tableDataPassword = $null
            
            if ($BackendSidecarMap -and $BackendSidecarMap.ContainsKey($sourcePath)) {
                $tableDataSource = $BackendSidecarMap[$sourcePath]
                $tableDataPassword = ""  # sidecars have no password
                Write-Status -Message "    Multi-backend: usando sidecar para '$sourcePath'" -Color Gray
            } elseif ($dataSource) {
                $tableDataSource = $dataSource
                $tableDataPassword = $dataPassword
            } else {
                Write-Status -Message "    ERROR: No hay fuente de datos para '$sourcePath'. Configurar BackendSidecarMap." -Color Red
                continue
            }

            # Abrir source (sidecar o backend) para obtener estructura y datos
            $srcEngine = New-DaoDbEngine
            $srcDb = $srcEngine.OpenDatabase($tableDataSource, $false, $false, $tableDataPassword)

            try {
                $srcTableDef = $srcDb.TableDefs[$sourceTableName]
                if (-not $srcTableDef) {
                    Write-Status -Message "    WARNING: Tabla '$sourceTableName' no encontrada en dataSource" -Color Yellow
                    continue
                }

                # Collect field definitions (BUG N FIX: preserve Attributes for AutoNumber/Required)
                $fieldDefs = @()
                for ($f = 0; $f -lt $srcTableDef.Fields.Count; $f++) {
                    $field = $srcTableDef.Fields[$f]
                    $fieldDefs += [pscustomobject]@{
                        Name = $field.Name
                        Type = $field.Type
                        Size = $field.Size
                        Attributes = $field.Attributes  # BUG N FIX: preserve AutoNumber, Required, etc.
                    }
                }

                # Eliminar la TableDef vinculada del sandbox
                $db.TableDefs.Delete($linked.TableName)

                # Crear nueva TableDef local en el sandbox
                $newTd = $db.CreateTableDef($linked.TableName)

                foreach ($fd in $fieldDefs) {
                    try {
                        $newField = $newTd.CreateField($fd.Name, $fd.Type, $fd.Size)
                        $newField.Attributes = $fd.Attributes  # BUG N FIX: restore Attributes
                        $newTd.Fields.Append($newField)
                    } catch {
                        Write-Status -Message "    WARNING: Campo '$($fd.Name)' no se pudo crear: $($_.Exception.Message)" -Color Yellow
                    }
                }

                $db.TableDefs.Append($newTd)

                # Copiar datos via recordset
                $srcRs = $srcDb.OpenRecordset($sourceTableName, 1)  # dbOpenTable
                $destRs = $db.OpenRecordset($linked.TableName, 2)    # dbOpenDynaset

                try {
                    while (-not $srcRs.EOF) {
                        $destRs.AddNew
                        for ($k = 0; $k -lt $srcRs.Fields.Count; $k++) {
                            try { $destRs.Fields[$k].Value = $srcRs.Fields[$k].Value } catch {}
                        }
                        $destRs.Update
                        $srcRs.MoveNext
                    }
                    Write-Status -Message "    OK $($linked.TableName) localizeada" -Color Gray
                } finally {
                    # BUG L FIX: Separate try-catch blocks ensure BOTH recordsets
                    # are closed even if one throws. Empty catch suppresses errors.
                    try { $srcRs.Close() } catch { }
                    try { $destRs.Close() } catch { }
                }

            } finally {
                # .Close() antes de FinalReleaseComObject — necesario para flush de writes y liberar file lock
                if ($srcDb) { try { $srcDb.Close() } catch {} }
                Close-DaoObjects -Objects @($srcDb, $srcEngine)
            }
        }

        # Compactar la BD para limpiar espacio — operacion puramente DAO
        Write-Status -Message "Compactando sandbox..." -Color Cyan
        try {
            # Cerrar DAO primero
            if ($db) { try { $db.Close() } catch {} }
            Close-DaoObjects -Objects @($db, $daoEngine)
            $daoEngine = $null
            $db = $null

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            # Compactar via DAO
            # BUG 6 FIX: CompactDatabase signature is (srcDb, destDb, destLocale, options, password)
            # destLocale="" for dbLangGeneral, options=128 for dbVersion120 (Access 2007+ .accdb)
            $compactEngine = New-DaoDbEngine
            $compactPwd = if ($SandboxPassword) { ";PWD=$SandboxPassword" } else { "" }
            $compactEngine.CompactDatabase($SandboxPath, $SandboxPath + ".tmp", "", 128, $compactPwd)
            Close-DaoObjects -Objects @($compactEngine)

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            Move-Item -Path ($SandboxPath + ".tmp") -Destination $SandboxPath -Force

        } catch {
            Write-Status -Message "WARNING: Compact falló: $($_.Exception.Message)" -Color Yellow
            # BUG E/J FIX: Clean up orphaned .tmp file if Move-Item failed
            Remove-Item -Path ($SandboxPath + ".tmp") -ErrorAction SilentlyContinue
        }

        Write-Status -Message "OK Sandbox localizeado exitosamente" -Color Green

    } finally {
        # .Close() antes de FinalReleaseComObject — necesario para flush de writes y liberar file lock antes del compact
        if ($db) { try { $db.Close() } catch {} }
        Close-DaoObjects -Objects @($db, $daoEngine)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# ============================================================
# VALIDATION
# ============================================================

if (-not (Test-DaoAvailable)) {
    throw "CRITICAL: DAO.DBEngine no disponible. Verificar que Microsoft Access o Access Database Engine esté instalado."
}

# ============================================================
# DISPATCH
# ============================================================

try {
    switch ($Action) {
        "New-Sandbox" {
            if ([string]::IsNullOrWhiteSpace($SourceBackend)) { throw "New-Sandbox requiere -SourceBackend" }
            if ([string]::IsNullOrWhiteSpace($SandboxPath)) { throw "New-Sandbox requiere -SandboxPath" }
            New-Sandbox -SourceBackend $SourceBackend -SandboxPath $SandboxPath
        }

        "Discover-LinkedTables" {
            if ([string]::IsNullOrWhiteSpace($SandboxPath)) { throw "Discover-LinkedTables requiere -SandboxPath" }
            $result = Discover-LinkedTables -SandboxPath $SandboxPath -Password $Password
            # Return via pipeline for capture
            $result | ForEach-Object { $_ }
        }

        "Make-Sidecar" {
            if ([string]::IsNullOrWhiteSpace($SourceBackend)) { throw "Make-Sidecar requiere -SourceBackend" }
            $sidecar = Make-Sidecar -SourceBackend $SourceBackend -Password $Password -SidecarSuffix $SidecarSuffix
            Write-Output $sidecar
        }

        "Localize-Sandbox" {
            if ([string]::IsNullOrWhiteSpace($SandboxPath)) { throw "Localize-Sandbox requiere -SandboxPath" }
            
            # Parse JSON backend sidecar map if provided
            $backendSidecarMap = $null
            if (-not [string]::IsNullOrWhiteSpace($BackendSidecarMapJson)) {
                try {
                    $backendSidecarMap = ConvertFrom-Json -InputObject $BackendSidecarMapJson -AsHashtable
                } catch {
                    throw "BackendSidecarMapJson inválido: $($_.Exception.Message)"
                }
            }
            
            if ([string]::IsNullOrWhiteSpace($SourceSidecar) -and [string]::IsNullOrWhiteSpace($SourceBackend) -and $null -eq $backendSidecarMap) {
                throw "Localize-Sandbox requiere -SourceSidecar, -SourceBackend, o -BackendSidecarMap"
            }
            Localize-Sandbox -SandboxPath $SandboxPath -SourceSidecar $SourceSidecar -SourceBackend $SourceBackend -Password $Password -SandboxPassword $SandboxPassword -BackendSidecarMap $backendSidecarMap
        }
    }
} catch {
    Write-Status -Message "ERROR: $($_.Exception.Message)" -Color Red
    throw
}
