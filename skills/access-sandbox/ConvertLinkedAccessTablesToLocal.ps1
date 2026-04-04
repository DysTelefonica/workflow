[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [string]$FrontendPath,

    [Parameter()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "FrontendPassword", Justification = "Requerido por especificacion del proyecto.")]
    [string]$FrontendPassword = "",

    [Parameter()]
    [switch]$KeepCopiedBackends,

    [Parameter()]
    [switch]$CleanPreviousSidecars,

    [Parameter()]
    [string]$BackupFolder = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

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

function New-DaoDbEngine {
    [CmdletBinding()]
    Param()

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

    return $null
}

function Resolve-PathSafe {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$PathValue
    )

    try {
        return [System.IO.Path]::GetFullPath($PathValue)
    } catch {
        return $PathValue
    }
}

function Escape-AccessName {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$Name
    )

    return "[" + ($Name -replace "\]", "]]" ) + "]"
}

function Test-DaoAvailable {
    [CmdletBinding()]
    Param()

    $engine = $null
    try {
        $engine = New-DaoDbEngine
        return ($null -ne $engine)
    } finally {
        if ($engine) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null } catch {} }
    }
}

function Get-AllowBypassKeyState {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password
    )

    $dbEngine = $null
    $database = $null
    $prop = $null

    try {
        $dbEngine = New-DaoDbEngine
        if (-not $dbEngine) { return $null }

        $connect = ""
        if (-not [string]::IsNullOrEmpty($Password)) {
            $connect = ";PWD=$Password"
        }

        try {
            $database = $dbEngine.OpenDatabase($AccessPath, $false, $false, $connect)
        } catch {
            return $null
        }

        try {
            $prop = $database.Properties("AllowBypassKey")
            return [pscustomobject]@{ Existed = $true; Value = [bool]$prop.Value }
        } catch {
            return [pscustomobject]@{ Existed = $false; Value = $null }
        }
    } finally {
        if ($database) { try { $database.Close() } catch {} }
        foreach ($obj in @($prop, $database, $dbEngine)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Enable-AllowBypassKey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password
    )

    $dbEngine = $null
    $database = $null
    $prop = $null
    $newProp = $null

    try {
        $dbEngine = New-DaoDbEngine
        if (-not $dbEngine) { return $false }

        $connect = ""
        if (-not [string]::IsNullOrEmpty($Password)) {
            $connect = ";PWD=$Password"
        }

        try {
            $database = $dbEngine.OpenDatabase($AccessPath, $false, $false, $connect)
        } catch {
            return $false
        }

        try {
            $prop = $database.Properties("AllowBypassKey")
            $prop.Value = $true
        } catch {
            $newProp = $database.CreateProperty("AllowBypassKey", 1, $true)
            $database.Properties.Append($newProp)
        }
        return $true
    } finally {
        if ($database) { try { $database.Close() } catch {} }
        foreach ($obj in @($newProp, $prop, $database, $dbEngine)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Restore-AllowBypassKey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password,
        $OriginalState
    )

    if (-not $OriginalState) { return }

    $dbEngine = $null
    $database = $null
    $prop = $null

    try {
        $dbEngine = New-DaoDbEngine
        if (-not $dbEngine) { return }

        $connect = ""
        if (-not [string]::IsNullOrEmpty($Password)) {
            $connect = ";PWD=$Password"
        }

        try {
            $database = $dbEngine.OpenDatabase($AccessPath, $false, $false, $connect)
        } catch {
            return
        }

        if ($OriginalState.Existed) {
            $prop = $database.Properties("AllowBypassKey")
            $prop.Value = [bool]$OriginalState.Value
        } else {
            try { $database.Properties.Delete("AllowBypassKey") } catch {}
        }
    } finally {
        if ($database) { try { $database.Close() } catch {} }
        foreach ($obj in @($prop, $database, $dbEngine)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Disable-StartupFeatures {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password
    )

    $dbEngine = New-DaoDbEngine
    if (-not $dbEngine) { return $null }

    $db = $null
    $restoreInfo = [ordered]@{
        RenamedAutoExec     = $false
        OriginalStartupForm = $null
        HasStartupForm      = $false
    }

    try {
        $connect = if ($Password) { ";PWD=$Password" } else { "" }
        $db = $dbEngine.OpenDatabase($AccessPath, $false, $false, $connect)

        try {
            $scripts = $db.Containers("Scripts")

            foreach ($doc in $scripts.Documents) {
                if ($doc.Name -eq "AutoExec_TraeBackup") {
                    $autoExecExists = $false
                    foreach ($d in $scripts.Documents) {
                        if ($d.Name -eq "AutoExec") { $autoExecExists = $true }
                    }
                    if (-not $autoExecExists) { $doc.Name = "AutoExec" }
                }
            }

            foreach ($doc in $scripts.Documents) {
                if ($doc.Name -eq "AutoExec") {
                    $doc.Name = "AutoExec_TraeBackup"
                    $restoreInfo.RenamedAutoExec = $true
                    break
                }
            }
        } catch {}

        try {
            $prop = $db.Properties("StartupForm")
            $restoreInfo.OriginalStartupForm = $prop.Value
            $restoreInfo.HasStartupForm = $true
            $db.Properties.Delete("StartupForm")
        } catch {}

        return [pscustomobject]$restoreInfo
    } catch {
        return $null
    } finally {
        if ($db) { try { $db.Close() } catch {} }
        if ($db) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } catch {} }
        if ($dbEngine) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dbEngine) | Out-Null } catch {} }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Restore-StartupFeatures {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password,
        $RestoreInfo
    )

    if (-not $RestoreInfo) { return }

    $dbEngine = $null
    $db = $null
    $newProp = $null

    try {
        $dbEngine = New-DaoDbEngine
        if (-not $dbEngine) { return }

        $connect = if ($Password) { ";PWD=$Password" } else { "" }
        $db = $dbEngine.OpenDatabase($AccessPath, $false, $false, $connect)

        if ($RestoreInfo.RenamedAutoExec) {
            try {
                $scripts = $db.Containers("Scripts")
                foreach ($doc in $scripts.Documents) {
                    if ($doc.Name -eq "AutoExec_TraeBackup") {
                        $doc.Name = "AutoExec"
                        break
                    }
                }
            } catch {}
        }

        if ($RestoreInfo.HasStartupForm) {
            try {
                $newProp = $db.CreateProperty("StartupForm", 10, $RestoreInfo.OriginalStartupForm)
                $db.Properties.Append($newProp)
            } catch {}
        }
    } catch {
    } finally {
        if ($db) { try { $db.Close() } catch {} }
        foreach ($obj in @($newProp, $db, $dbEngine)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-ProcessIdFromHwnd {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][IntPtr]$Hwnd
    )

    if (-not ([System.Management.Automation.PSTypeName]"Win32.NativeMethods").Type) {
        Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
"@
    }

    [uint32]$pid = 0
    [Win32.NativeMethods]::GetWindowThreadProcessId($Hwnd, [ref]$pid) | Out-Null
    return [int]$pid
}

function Open-AccessDatabase {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password
    )

    $access = $null
    $originalBypass = $null
    $accessPid = $null
    $prePids = @()
    $startupInfo = $null

    try {
        $originalBypass = Get-AllowBypassKeyState -AccessPath $AccessPath -Password $Password

        $bypassOk = Enable-AllowBypassKey -AccessPath $AccessPath -Password $Password
        if (-not $bypassOk) {
            Write-Status -Message "ADVERTENCIA: No se pudo habilitar AllowBypassKey; abriendo de todas formas." -Color Yellow
        }

        $startupInfo = Disable-StartupFeatures -AccessPath $AccessPath -Password $Password
        if (-not $startupInfo) {
            throw "CRITICAL: No se pudo deshabilitar AutoExec/StartupForm. Se aborta la apertura para evitar ejecucion no desatendida."
        }

        try {
            $prePids = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        } catch {
            $prePids = @()
        }

        $access = New-Object -ComObject Access.Application
        $access.Visible = $false
        $access.UserControl = $false
        $access.AutomationSecurity = 1

        try {
            $hwnd = [IntPtr]$access.hWndAccessApp
            if ($hwnd -and $hwnd -ne [IntPtr]::Zero) {
                $accessPid = Get-ProcessIdFromHwnd -Hwnd $hwnd
            }
        } catch {}

        $access.OpenCurrentDatabase($AccessPath, $false, $Password)
        try { $access.DoCmd.SetWarnings($false) } catch {}

        try {
            if (-not $accessPid) {
                $hwnd2 = [IntPtr]$access.hWndAccessApp
                if ($hwnd2 -and $hwnd2 -ne [IntPtr]::Zero) {
                    $accessPid = Get-ProcessIdFromHwnd -Hwnd $hwnd2
                }
            }
        } catch {}

        try {
            $post = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -Property Id, StartTime)
            $new = @($post | Where-Object { $_.Id -notin $prePids })
            if ($new.Count -ge 1) {
                $picked = $new | Sort-Object -Property StartTime -Descending | Select-Object -First 1
                $accessPid = [int]$picked.Id
            }
        } catch {}

        return [pscustomobject]@{
            AccessApplication = $access
            OriginalBypass    = $originalBypass
            StartupInfo       = $startupInfo
            ProcessId         = $accessPid
        }
    } catch {
        if ($access) {
            try { $access.CloseCurrentDatabase() } catch {}
            try { $access.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) | Out-Null } catch {}
        }

        if ($originalBypass) {
            try { Restore-AllowBypassKey -AccessPath $AccessPath -Password $Password -OriginalState $originalBypass } catch {}
        }
        if ($startupInfo) {
            try { Restore-StartupFeatures -AccessPath $AccessPath -Password $Password -RestoreInfo $startupInfo } catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()

        throw
    }
}

function Close-AccessDatabase {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password
    )

    $access = $Session.AccessApplication
    $orig = $Session.OriginalBypass
    $startupInfo = $Session.StartupInfo
    $accessPid = $Session.ProcessId

    if ($access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
    }

    if ($Session.AccessApplication) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Session.AccessApplication) | Out-Null } catch {}
    }

    try { Restore-AllowBypassKey -AccessPath $AccessPath -Password $Password -OriginalState $orig } catch {}
    try { Restore-StartupFeatures -AccessPath $AccessPath -Password $Password -RestoreInfo $startupInfo } catch {}

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($accessPid) {
        try { Stop-Process -Id $accessPid -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Get-ConnectValue {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ConnectString,
        [Parameter(Mandatory = $true)][string]$Key
    )

    $m = [regex]::Match($ConnectString, ('(?i)(?:^|;){0}=([^;]*)' -f [regex]::Escape($Key)))
    if ($m.Success) {
        return $m.Groups[1].Value
    }
    return $null
}

function Get-ConnectDatabasePath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ConnectString
    )

    $dbPath = Get-ConnectValue -ConnectString $ConnectString -Key 'DATABASE'
    if ([string]::IsNullOrWhiteSpace($dbPath)) { return $null }
    return (Resolve-PathSafe -PathValue $dbPath)
}

function Set-ConnectDatabasePath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ConnectString,
        [Parameter(Mandatory = $true)][string]$NewDatabasePath
    )

    if ([string]::IsNullOrWhiteSpace($ConnectString)) {
        throw "ConnectString vacio."
    }

    if ($ConnectString -match '(?i)(^|;)DATABASE=') {
        return [regex]::Replace(
            $ConnectString,
            '(?i)(^|;)DATABASE=([^;]+)',
            ('$1DATABASE=' + $NewDatabasePath)
        )
    }

    if ($ConnectString.EndsWith(';')) {
        return $ConnectString + "DATABASE=$NewDatabasePath"
    }

    return $ConnectString + ";DATABASE=$NewDatabasePath"
}

function Get-DaoOpenConnectFromTableConnect {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ConnectString
    )

    $pwd = Get-ConnectValue -ConnectString $ConnectString -Key 'PWD'
    if ([string]::IsNullOrEmpty($pwd)) {
        return ""
    }

    return ";PWD=$pwd"
}

function Get-ExistingSidecarPaths {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$FrontendPath
    )

    $frontendDir = Split-Path -LiteralPath $FrontendPath -Parent
    $frontendName = [System.IO.Path]::GetFileNameWithoutExtension($FrontendPath)
    $frontendExt = [System.IO.Path]::GetExtension($FrontendPath)
    $pattern = "{0}__sidecar__*{1}" -f $frontendName, $frontendExt

    return @(Get-ChildItem -LiteralPath $frontendDir -Filter $pattern -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
}

function New-BackupPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$FrontendPath,
        [string]$BackupFolder
    )

    $frontendName = [System.IO.Path]::GetFileNameWithoutExtension($FrontendPath)
    $frontendExt = [System.IO.Path]::GetExtension($FrontendPath)
    $folder = if ([string]::IsNullOrWhiteSpace($BackupFolder)) {
        Split-Path -LiteralPath $FrontendPath -Parent
    } else {
        Resolve-PathSafe -PathValue $BackupFolder
    }

    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    return (Join-Path $folder ("{0}__backup__{1}{2}" -f $frontendName, $stamp, $frontendExt))
}

function New-SidecarPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$FrontendPath,
        [Parameter(Mandatory = $true)][string]$BackendOriginalPath
    )

    $frontendDir = Split-Path -LiteralPath $FrontendPath -Parent
    $frontendName = [System.IO.Path]::GetFileNameWithoutExtension($FrontendPath)
    $backendBaseName = [System.IO.Path]::GetFileNameWithoutExtension($BackendOriginalPath)
    $backendExt = [System.IO.Path]::GetExtension($BackendOriginalPath)

    $candidate = Join-Path $frontendDir ("{0}__sidecar__{1}{2}" -f $frontendName, $backendBaseName, $backendExt)

    if ((Resolve-PathSafe -PathValue $candidate).ToLowerInvariant() -eq (Resolve-PathSafe -PathValue $BackendOriginalPath).ToLowerInvariant()) {
        $candidate = Join-Path $frontendDir ("{0}__sidecar__{1}_{2}{3}" -f $frontendName, $backendBaseName, ([guid]::NewGuid().ToString('N')), $backendExt)
    }

    return $candidate
}

function Open-DaoDatabaseFromTableConnect {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$DatabasePath,
        [Parameter(Mandatory = $true)][string]$TableConnect,
        [bool]$ReadOnly = $true
    )

    $dbEngine = New-DaoDbEngine
    if (-not $dbEngine) {
        throw "No se pudo crear DAO.DBEngine."
    }

    $daoConnect = Get-DaoOpenConnectFromTableConnect -ConnectString $TableConnect
    $db = $null

    try {
        $db = $dbEngine.OpenDatabase($DatabasePath, $false, $ReadOnly, $daoConnect)
        return [pscustomobject]@{
            Engine = $dbEngine
            Database = $db
        }
    } catch {
        if ($db) { try { $db.Close() } catch {} }
        if ($db) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } catch {} }
        if ($dbEngine) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dbEngine) | Out-Null } catch {} }
        throw
    }
}

function Close-DaoDatabaseHandle {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Handle
    )

    if ($Handle.Database) { try { $Handle.Database.Close() } catch {} }
    if ($Handle.Database) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Handle.Database) | Out-Null } catch {} }
    if ($Handle.Engine) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Handle.Engine) | Out-Null } catch {} }
}

function Get-LinkedAccessTablesGroupedByBackend {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Db
    )

    $result = @{}

    foreach ($td in $Db.TableDefs) {
        $tableName = ''
        $connect = ''
        $source = ''

        try { $tableName = [string]$td.Name } catch {}
        try { $connect = [string]$td.Connect } catch {}
        try { $source = [string]$td.SourceTableName } catch {}

        if ([string]::IsNullOrWhiteSpace($tableName)) { continue }
        if ($tableName.StartsWith('MSys', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if ([string]::IsNullOrWhiteSpace($connect)) { continue }
        if ($connect -notmatch '(?i)^\s*MS\s*Access\s*;') { continue }

        $backendPath = Get-ConnectDatabasePath -ConnectString $connect
        if ([string]::IsNullOrWhiteSpace($backendPath)) { continue }

        if (-not $result.ContainsKey($backendPath)) {
            $result[$backendPath] = New-Object System.Collections.ArrayList
        }

        [void]$result[$backendPath].Add([pscustomobject]@{
            LocalTableName  = $tableName
            SourceTableName = if ([string]::IsNullOrWhiteSpace($source)) { $tableName } else { $source }
            OriginalConnect = $connect
        })
    }

    return $result
}

function Test-BackendAccessible {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$BackendPath,
        [Parameter(Mandatory = $true)][string]$OriginalConnect
    )

    $handle = $null
    try {
        $handle = Open-DaoDatabaseFromTableConnect -DatabasePath $BackendPath -TableConnect $OriginalConnect -ReadOnly $true
        return $true
    } catch {
        return $false
    } finally {
        if ($handle) { Close-DaoDatabaseHandle -Handle $handle }
    }
}

function Remove-TableDefIfExists {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Db,
        [Parameter(Mandatory = $true)][string]$TableName
    )

    try {
        $null = $Db.TableDefs.Item($TableName)
        $Db.TableDefs.Delete($TableName)
        $Db.TableDefs.Refresh()
    } catch {}
}

function New-TempName {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$Prefix
    )

    return ('{0}{1}' -f $Prefix, [guid]::NewGuid().ToString('N'))
}

function Get-FieldDescription {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Field
    )

    try {
        return [string]$Field.Properties('Description').Value
    } catch {
        return $null
    }
}

function Copy-DescriptionProperty {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Field,
        [string]$Description
    )

    if ([string]::IsNullOrEmpty($Description)) { return }

    try {
        $prop = $Field.CreateProperty('Description', 10, $Description)
        $Field.Properties.Append($prop)
    } catch {
        try { $Field.Properties('Description').Value = $Description } catch {}
    }
}

function New-LocalTableFromSourceDefinition {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$FrontendDb,
        [Parameter(Mandatory = $true)]$SourceTableDef,
        [Parameter(Mandatory = $true)][string]$TempLocalTableName
    )

    $newTable = $null
    $fieldDescriptions = @{}
    $appendedTable = $null

    try {
        $newTable = $FrontendDb.CreateTableDef($TempLocalTableName)

        foreach ($sourceField in $SourceTableDef.Fields) {
            $newField = $null
            try {
                $fieldName = [string]$sourceField.Name
                $fieldType = [int]$sourceField.Type
                $fieldSize = 0
                try { $fieldSize = [int]$sourceField.Size } catch { $fieldSize = 0 }

                if ($fieldSize -gt 0) {
                    $newField = $newTable.CreateField($fieldName, $fieldType, $fieldSize)
                } else {
                    $newField = $newTable.CreateField($fieldName, $fieldType)
                }

                try { $newField.Attributes = $sourceField.Attributes } catch {}
                try { $newField.Required = $sourceField.Required } catch {}
                try { $newField.AllowZeroLength = $sourceField.AllowZeroLength } catch {}
                try { if ($null -ne $sourceField.DefaultValue -and "$($sourceField.DefaultValue)" -ne '') { $newField.DefaultValue = $sourceField.DefaultValue } } catch {}
                try { if ($null -ne $sourceField.ValidationRule -and "$($sourceField.ValidationRule)" -ne '') { $newField.ValidationRule = $sourceField.ValidationRule } } catch {}
                try { if ($null -ne $sourceField.ValidationText -and "$($sourceField.ValidationText)" -ne '') { $newField.ValidationText = $sourceField.ValidationText } } catch {}

                $newTable.Fields.Append($newField)
                $desc = Get-FieldDescription -Field $sourceField
                if (-not [string]::IsNullOrEmpty($desc)) {
                    $fieldDescriptions[$fieldName] = $desc
                }
            } finally {
                if ($newField) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($newField) | Out-Null } catch {} }
                if ($sourceField) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sourceField) | Out-Null } catch {} }
            }
        }

        $FrontendDb.TableDefs.Append($newTable)
        $FrontendDb.TableDefs.Refresh()

        $appendedTable = $FrontendDb.TableDefs.Item($TempLocalTableName)
        foreach ($fieldName in $fieldDescriptions.Keys) {
            $appendedField = $null
            try {
                $appendedField = $appendedTable.Fields.Item([string]$fieldName)
                Copy-DescriptionProperty -Field $appendedField -Description ([string]$fieldDescriptions[$fieldName])
            } finally {
                if ($appendedField) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($appendedField) | Out-Null } catch {} }
            }
        }
    } finally {
        if ($appendedTable) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($appendedTable) | Out-Null } catch {} }
        if ($newTable) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($newTable) | Out-Null } catch {} }
    }
}

function New-TemporaryLinkedTable {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$FrontendDb,
        [Parameter(Mandatory = $true)][string]$TempLinkedTableName,
        [Parameter(Mandatory = $true)][string]$SourceTableName,
        [Parameter(Mandatory = $true)][string]$OriginalConnect,
        [Parameter(Mandatory = $true)][string]$SidecarPath
    )

    $tmpTdf = $null
    try {
        $tmpTdf = $FrontendDb.CreateTableDef($TempLinkedTableName)
        $tmpTdf.SourceTableName = $SourceTableName
        $tmpTdf.Connect = Set-ConnectDatabasePath -ConnectString $OriginalConnect -NewDatabasePath $SidecarPath
        $FrontendDb.TableDefs.Append($tmpTdf)
        $FrontendDb.TableDefs.Refresh()
        $null = $FrontendDb.TableDefs.Item($TempLinkedTableName).Fields.Count
    } finally {
        if ($tmpTdf) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($tmpTdf) | Out-Null } catch {} }
    }
}

function Get-InsertableFieldNames {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$SourceTableDef
    )

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($field in $SourceTableDef.Fields) {
        $skip = $false
        try {
            if (($field.Attributes -band 32768) -ne 0) { $skip = $true } # dbSystemField
        } catch {}

        if (-not $skip) {
            [void]$names.Add([string]$field.Name)
        }
    }
    return @($names)
}

function Insert-TableDataFromLinkedTable {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$FrontendDb,
        [Parameter(Mandatory = $true)][string]$TempLocalTableName,
        [Parameter(Mandatory = $true)][string]$TempLinkedTableName,
        [Parameter(Mandatory = $true)][string[]]$FieldNames
    )

    if (-not $FieldNames -or $FieldNames.Count -eq 0) {
        return
    }

    $columnList = ($FieldNames | ForEach-Object { Escape-AccessName -Name $_ }) -join ', '
    $sql = "INSERT INTO {0} ({1}) SELECT {1} FROM {2}" -f `
        (Escape-AccessName -Name $TempLocalTableName), `
        $columnList, `
        (Escape-AccessName -Name $TempLinkedTableName)

    $FrontendDb.Execute($sql, 128)
}

function Get-TableRowCount {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Db,
        [Parameter(Mandatory = $true)][string]$TableName
    )

    $rs = $null
    try {
        $sql = "SELECT COUNT(*) AS Cnt FROM {0}" -f (Escape-AccessName -Name $TableName)
        $rs = $Db.OpenRecordset($sql)
        return [int]$rs.Fields('Cnt').Value
    } finally {
        if ($rs) { try { $rs.Close() } catch {} }
        if ($rs) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($rs) | Out-Null } catch {} }
    }
}

function Recreate-Indexes {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$FrontendDb,
        [Parameter(Mandatory = $true)]$SourceTableDef,
        [Parameter(Mandatory = $true)][string]$TempLocalTableName
    )

    $targetTable = $null
    try {
        $targetTable = $FrontendDb.TableDefs.Item($TempLocalTableName)

        foreach ($sourceIndex in $SourceTableDef.Indexes) {
            $newIndex = $null
            try {
                $indexName = [string]$sourceIndex.Name
                if ([string]::IsNullOrWhiteSpace($indexName)) {
                    $indexName = New-TempName -Prefix '__idx__'
                }

                $newIndex = $targetTable.CreateIndex($indexName)
                try { $newIndex.Primary = $sourceIndex.Primary } catch {}
                try { $newIndex.Unique = $sourceIndex.Unique } catch {}
                try { $newIndex.Required = $sourceIndex.Required } catch {}
                try { $newIndex.IgnoreNulls = $sourceIndex.IgnoreNulls } catch {}
                try { $newIndex.Clustered = $sourceIndex.Clustered } catch {}

                foreach ($sourceIndexField in $sourceIndex.Fields) {
                    $newIndexField = $null
                    try {
                        $newIndexField = $newIndex.CreateField([string]$sourceIndexField.Name)
                        try { $newIndexField.Attributes = $sourceIndexField.Attributes } catch {}
                        $newIndex.Fields.Append($newIndexField)
                    } finally {
                        if ($newIndexField) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($newIndexField) | Out-Null } catch {} }
                    }
                }

                $targetTable.Indexes.Append($newIndex)
            } finally {
                if ($newIndex) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($newIndex) | Out-Null } catch {} }
            }
        }
    } finally {
        if ($targetTable) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($targetTable) | Out-Null } catch {} }
    }
}

function Rename-Table {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$AccessApp,
        [Parameter(Mandatory = $true)][string]$OldName,
        [Parameter(Mandatory = $true)][string]$NewName
    )

    $AccessApp.DoCmd.Rename($NewName, 0, $OldName)
}

function Convert-LinkedTableToLocal {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$AccessApp,
        [Parameter(Mandatory = $true)]$FrontendDb,
        [Parameter(Mandatory = $true)][string]$LocalTableName,
        [Parameter(Mandatory = $true)][string]$SourceTableName,
        [Parameter(Mandatory = $true)][string]$OriginalConnect,
        [Parameter(Mandatory = $true)][string]$SidecarPath
    )

    $backendHandle = $null
    $sourceTableDef = $null
    $tempLocalTableName = New-TempName -Prefix '__tmp_localcopy__'
    $tempLinkedTableName = New-TempName -Prefix '__tmp_link__'
    $tempLinkedCreated = $false
    $tempLocalCreated = $false

    try {
        $backendHandle = Open-DaoDatabaseFromTableConnect -DatabasePath $SidecarPath -TableConnect $OriginalConnect -ReadOnly $true
        $sourceTableDef = $backendHandle.Database.TableDefs.Item($SourceTableName)

        New-LocalTableFromSourceDefinition -FrontendDb $FrontendDb -SourceTableDef $sourceTableDef -TempLocalTableName $tempLocalTableName
        $tempLocalCreated = $true

        New-TemporaryLinkedTable -FrontendDb $FrontendDb -TempLinkedTableName $tempLinkedTableName -SourceTableName $SourceTableName -OriginalConnect $OriginalConnect -SidecarPath $SidecarPath
        $tempLinkedCreated = $true

        $fieldNames = Get-InsertableFieldNames -SourceTableDef $sourceTableDef
        Insert-TableDataFromLinkedTable -FrontendDb $FrontendDb -TempLocalTableName $tempLocalTableName -TempLinkedTableName $tempLinkedTableName -FieldNames $fieldNames

        $sourceCount = Get-TableRowCount -Db $FrontendDb -TableName $tempLinkedTableName
        $localCount = Get-TableRowCount -Db $FrontendDb -TableName $tempLocalTableName
        if ($sourceCount -ne $localCount) {
            throw "El numero de registros no coincide para '$LocalTableName'. Origen=$sourceCount, Destino=$localCount"
        }

        Recreate-Indexes -FrontendDb $FrontendDb -SourceTableDef $sourceTableDef -TempLocalTableName $tempLocalTableName

        Remove-TableDefIfExists -Db $FrontendDb -TableName $LocalTableName
        Rename-Table -AccessApp $AccessApp -OldName $tempLocalTableName -NewName $LocalTableName
        $tempLocalCreated = $false

        Remove-TableDefIfExists -Db $FrontendDb -TableName $tempLinkedTableName
        $tempLinkedCreated = $false
    }
    catch {
        throw "Fallo procesando la tabla '$LocalTableName' desde '$SourceTableName'. Error: $($_.Exception.Message)"
    }
    finally {
        if ($tempLinkedCreated) {
            try { Remove-TableDefIfExists -Db $FrontendDb -TableName $tempLinkedTableName } catch {}
        }
        if ($tempLocalCreated) {
            try { Remove-TableDefIfExists -Db $FrontendDb -TableName $tempLocalTableName } catch {}
        }
        if ($sourceTableDef) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sourceTableDef) | Out-Null } catch {} }
        if ($backendHandle) { Close-DaoDatabaseHandle -Handle $backendHandle }
    }
}

$session = $null
$db = $null
$linkedGroups = $null
$sidecarsCreated = @()
$sidecarMap = @{}
$backupPath = $null
$success = $false

try {
    $FrontendPath = Resolve-PathSafe -PathValue $FrontendPath

    if (-not (Test-Path -LiteralPath $FrontendPath)) {
        throw "No existe el frontend: $FrontendPath"
    }

    if (-not (Test-DaoAvailable)) {
        throw "DAO no esta disponible en esta maquina."
    }

    $previousSidecars = @(Get-ExistingSidecarPaths -FrontendPath $FrontendPath)
    if ($previousSidecars.Count -gt 0) {
        if ($CleanPreviousSidecars) {
            foreach ($oldSidecar in $previousSidecars) {
                Write-Status -Message ("Eliminando sidecar previo: {0}" -f $oldSidecar) -Color Yellow
                Remove-Item -LiteralPath $oldSidecar -Force
            }
        } else {
            $joined = $previousSidecars -join [Environment]::NewLine
            throw "Se han encontrado sidecars previos. Ejecute con -CleanPreviousSidecars para limpiarlos antes de continuar:`n$joined"
        }
    }

    $backupPath = New-BackupPath -FrontendPath $FrontendPath -BackupFolder $BackupFolder
    Write-Status -Message ("Creando backup del frontend: {0}" -f $backupPath) -Color Cyan
    Copy-Item -LiteralPath $FrontendPath -Destination $backupPath -Force

    Write-Status -Message "Abriendo frontend en modo oculto..." -Color Cyan
    $session = Open-AccessDatabase -AccessPath $FrontendPath -Password $FrontendPassword
    $db = $session.AccessApplication.CurrentDb()

    Write-Status -Message "Detectando tablas vinculadas a otros Access..." -Color Cyan
    $linkedGroups = Get-LinkedAccessTablesGroupedByBackend -Db $db

    if (-not $linkedGroups -or $linkedGroups.Keys.Count -eq 0) {
        Write-Status -Message "No hay tablas vinculadas a bases Access. Nada que hacer." -Color Yellow
        $success = $true
        return
    }

    $tableCount = 0
    foreach ($k in $linkedGroups.Keys) { $tableCount += $linkedGroups[$k].Count }
    Write-Status -Message ("Backends Access distintos detectados: {0}" -f $linkedGroups.Keys.Count) -Color Green
    Write-Status -Message ("Tablas vinculadas detectadas: {0}" -f $tableCount) -Color Green

    foreach ($backendOriginal in $linkedGroups.Keys) {
        if (-not (Test-Path -LiteralPath $backendOriginal)) {
            throw "No existe el backend origen referenciado por una vinculacion: $backendOriginal"
        }

        $tables = $linkedGroups[$backendOriginal]
        $sampleConnect = [string]$tables[0].OriginalConnect

        if (-not (Test-BackendAccessible -BackendPath $backendOriginal -OriginalConnect $sampleConnect)) {
            throw "No se puede abrir el backend con la Connect original de sus vinculaciones: $backendOriginal"
        }

        $sidecarPath = New-SidecarPath -FrontendPath $FrontendPath -BackendOriginalPath $backendOriginal
        Write-Status -Message ("Copiando sidecar: {0} -> {1}" -f $backendOriginal, $sidecarPath) -Color DarkCyan
        Copy-Item -LiteralPath $backendOriginal -Destination $sidecarPath -Force

        if (-not (Test-BackendAccessible -BackendPath $sidecarPath -OriginalConnect $sampleConnect)) {
            throw "No se puede abrir el sidecar recien copiado: $sidecarPath"
        }

        $sidecarMap[$backendOriginal] = $sidecarPath
        $sidecarsCreated += $sidecarPath
    }

    foreach ($backendOriginal in $linkedGroups.Keys) {
        $sidecarPath = [string]$sidecarMap[$backendOriginal]
        $tables = $linkedGroups[$backendOriginal]

        Write-Status -Message ("Procesando backend: {0}" -f $backendOriginal) -Color Cyan

        foreach ($t in $tables) {
            $localTableName = [string]$t.LocalTableName
            $sourceTableName = [string]$t.SourceTableName
            $originalConnect = [string]$t.OriginalConnect

            Write-Status -Message ("  Convirtiendo tabla vinculada a local: {0}" -f $localTableName) -Color Gray

            Convert-LinkedTableToLocal `
                -AccessApp $session.AccessApplication `
                -FrontendDb $db `
                -LocalTableName $localTableName `
                -SourceTableName $sourceTableName `
                -OriginalConnect $originalConnect `
                -SidecarPath $sidecarPath
        }
    }

    try { $db.TableDefs.Refresh() } catch {}

    $success = $true
    Write-Status -Message "Proceso completado correctamente." -Color Green
}
catch {
    Write-Status -Message ("ERROR: {0}" -f $_.Exception.Message) -Color Red
    Write-Status -Message ("Backup disponible en: {0}" -f $backupPath) -Color Yellow
    throw
}
finally {
    if ($db) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null } catch {}
    }

    if ($session) {
        try { Close-AccessDatabase -Session $session -AccessPath $FrontendPath -Password $FrontendPassword } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($success) {
        if (-not $KeepCopiedBackends) {
            foreach ($p in $sidecarsCreated) {
                try {
                    if (Test-Path -LiteralPath $p) {
                        Remove-Item -LiteralPath $p -Force
                        Write-Status -Message ("Sidecar eliminado: {0}" -f $p) -Color DarkGray
                    }
                } catch {
                    Write-Status -Message ("ADVERTENCIA: No se pudo eliminar el sidecar: {0}" -f $p) -Color Yellow
                }
            }
        }
    } else {
        if ($sidecarsCreated.Count -gt 0) {
            Write-Status -Message "La ejecucion fallo; se conservan los sidecars para diagnostico." -Color Yellow
        }
    }
}
