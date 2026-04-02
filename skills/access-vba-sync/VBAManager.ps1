[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("Export", "Import", "Fix-Encoding", "Generate-ERD", "Delete", "Rename", "List")]
    [string]$Action,

    [Parameter()]
    [string]$AccessPath,

    [Parameter()]
    [Alias("DestinationPath")]
    [string]$DestinationRoot,

    # FIX: ModuleName con Position alto para que nunca compita posicionalmente con otros parametros.
    # Siempre pasar con nombre explicito: -ModuleName "A" "B" "C"
    # Evita que PowerShell asigne valores del array a -Location u otros parametros con ValidateSet.
    [Parameter(Position = 100)]
    [string[]]$ModuleName,

    # FIX: Location sin Position para que no participe en binding posicional automatico
    # y nunca compita con los valores del array de -ModuleName.
    [Parameter()]
    [ValidateSet("Both", "Src", "Access")]
    [string]$Location = "Both",

    [Parameter()]
    [ValidateSet("Auto", "Form", "Code")]
    [string]$ImportMode = "Auto",

    [Parameter()]
    [string]$BackendPath,

    [Parameter()]
    [string]$ErdPath,

    # NUEVO: nombre destino para Rename
    [Parameter()]
    [string]$NewModuleName,

    # NUEVO: switch para borrar tambien de src/ en Delete
    [Parameter()]
    [switch]$DeleteFromSrc,

    [Parameter()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "Password", Justification = "Requerido por especificacion del proyecto.")]
    [string]$Password = "dpddpd"
)

$ErrorActionPreference = "Stop"

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

function Get-AllowBypassKeyState {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
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
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Enable-AllowBypassKey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
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
            # 1 = dbBoolean, sin cast [int16] para evitar problemas COM
            $newProp = $database.CreateProperty("AllowBypassKey", 1, $true)
            $database.Properties.Append($newProp)
        }
        return $true
    } finally {
        if ($database) { try { $database.Close() } catch {} }
        foreach ($obj in @($newProp, $prop, $database, $dbEngine)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Restore-AllowBypassKey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
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
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
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
                    foreach ($d in $scripts.Documents) { if ($d.Name -eq "AutoExec") { $autoExecExists = $true } }
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
        if ($db) { $db.Close(); [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null }
        # FIX: liberar $dbEngine que antes quedaba vivo
        if ($dbEngine) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dbEngine) | Out-Null }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
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

    $dbEngine = New-DaoDbEngine
    $db = $null

    try {
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
                # 10 = dbText, sin cast [int16] para evitar problemas COM
                $newProp = $db.CreateProperty("StartupForm", 10, $RestoreInfo.OriginalStartupForm)
                $db.Properties.Append($newProp)
            } catch {}
        }
    } catch {
    } finally {
        if ($db) { $db.Close(); [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null }
        if ($dbEngine) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dbEngine) | Out-Null }
    }
}

function Resolve-AccessPath {
    [CmdletBinding()]
    Param(
        [string]$AccessPath
    )

    if (-not [string]::IsNullOrWhiteSpace($AccessPath)) {
        return (Resolve-Path -Path $AccessPath).Path
    }

    $candidates = Get-ChildItem -Path (Get-Location) -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @(".accdb", ".accde", ".mdb", ".mde") } |
        Sort-Object -Property Name

    if (-not $candidates -or $candidates.Count -eq 0) {
        throw "No se encontro ningun archivo .accdb/.accde/.mdb/.mde en el directorio actual."
    }

    if ($candidates.Count -gt 1) {
        Write-Status -Message "ADVERTENCIA: Se encontraron varias BDs; eligiendo determinista (alfabetico):" -Color Yellow
        foreach ($c in $candidates) { Write-Status -Message (" - {0}" -f $c.Name) -Color Yellow }
    }

    return $candidates[0].FullName
}

function Resolve-DestinationRoot {
    [CmdletBinding()]
    Param(
        [string]$DestinationRoot
    )

    if ([string]::IsNullOrWhiteSpace($DestinationRoot)) {
        $DestinationRoot = Join-Path -Path (Get-Location) -ChildPath "src"
    }

    if (-not (Test-Path -Path $DestinationRoot)) {
        New-Item -Path $DestinationRoot -ItemType Directory -Force | Out-Null
    }

    return (Resolve-Path -Path $DestinationRoot).Path
}

function Resolve-ModulesPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$DestinationRoot,
        [Parameter(Mandatory = $true)][string]$AccessPath,
        # FIX: ValidateSet ampliado con las nuevas acciones
        [Parameter(Mandatory = $true)][ValidateSet("Export", "Import", "Fix-Encoding", "Generate-ERD", "Delete", "Rename", "List")][string]$Action
    )
    if (-not (Test-Path -Path $DestinationRoot)) {
        if ($Action -in @("Export", "Fix-Encoding", "Delete", "Rename")) {
            New-Item -ItemType Directory -Force -Path $DestinationRoot | Out-Null
        } else {
            throw ("No existe la carpeta de modulos: {0}" -f $DestinationRoot)
        }
    }

    return (Resolve-Path -Path $DestinationRoot).Path
}

function Get-FileEncodingInfo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [pscustomobject]@{ HasUtf8Bom = $true; Bytes = $bytes }
    }
    return [pscustomobject]@{ HasUtf8Bom = $false; Bytes = $bytes }
}

function Write-Utf8NoBom {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $utf8NoBom)
}

function Convert-AnsiToUtf8NoBom {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    $ansi = [System.Text.Encoding]::GetEncoding(1252)
    $text = [System.IO.File]::ReadAllText($InputPath, $ansi)
    Write-Utf8NoBom -Path $OutputPath -Text $text
}

function Convert-Utf8ToAnsiTempFile {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$TempPath
    )

    $utf8 = [System.Text.Encoding]::UTF8
    $ansi = [System.Text.Encoding]::GetEncoding(1252)
    $text = [System.IO.File]::ReadAllText($InputPath, $utf8)
    [System.IO.File]::WriteAllText($TempPath, $text, $ansi)
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

function Open-AccessDatabase {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
        [string]$Password
    )

    $access = $null
    $originalBypass = $null
    $accessPid = $null
    $vbe = $null
    $vbProject = $null
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

        $vbe = $access.VBE
        $vbProject = $vbe.ActiveVBProject

        return [pscustomobject]@{
            AccessApplication = $access
            Vbe               = $vbe
            VbProject         = $vbProject
            OriginalBypass    = $originalBypass
            StartupInfo       = $startupInfo
            ProcessId         = $accessPid
        }
    } catch {
        if ($access) {
            try { $access.Quit() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($access) | Out-Null } catch {}
        }
        foreach ($obj in @($vbProject, $vbe)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
        if ($originalBypass) {
            try { Restore-AllowBypassKey -AccessPath $AccessPath -Password $Password -OriginalState $originalBypass } catch {}
        }
        if ($startupInfo) {
            try { Restore-StartupFeatures -AccessPath $AccessPath -Password $Password -RestoreInfo $startupInfo } catch {}
        }
        throw
    }
}

function Close-AccessDatabase {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
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

    foreach ($obj in @($Session.VbProject, $Session.Vbe, $Session.AccessApplication)) {
        if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
    }

    try { Restore-AllowBypassKey -AccessPath $AccessPath -Password $Password -OriginalState $orig } catch {}
    try { Restore-StartupFeatures -AccessPath $AccessPath -Password $Password -RestoreInfo $startupInfo } catch {}

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if ($accessPid) {
        try { Stop-Process -Id $accessPid -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Get-ComponentFolder {
    Param([Parameter(Mandatory = $true)]$Component, [string]$ModuleName)
    $name = if ($ModuleName) { $ModuleName } else { $Component.Name }
    if ($name -match "^Form_|^frm") { return "forms" }
    $t = $Component.Type
    if ($t -eq 1) { return "modules" }
    if ($t -eq 2) { return "classes" }
    if ($t -eq 100) { return "forms" }  # FIX: vbext_ct_Document es formulario, no clase
    if ($t -eq 3) { return "forms" }
    return $null
}

function Get-ComponentExtension {
    Param([Parameter(Mandatory = $true)]$Component, [string]$ModuleName)
    $name = if ($ModuleName) { $ModuleName } else { $Component.Name }
    if ($name -match "^Form_|^frm") { return ".form.txt" }
    $t = $Component.Type
    if ($t -eq 1) { return ".bas" }
    if ($t -eq 2) { return ".cls" }
    if ($t -eq 100) { return ".form.txt" }
    if ($t -eq 3) { return ".form.txt" }
    return $null
}

# NUEVO: obtener nombre legible del tipo de componente
function Get-ComponentTypeName {
    Param([Parameter(Mandatory = $true)][int]$Type)
    switch ($Type) {
        1   { return "Module" }
        2   { return "Class" }
        3   { return "Form" }
        100 { return "Form" }
        default { return "Type$Type" }
    }
}

function Export-VbaModule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        $AccessApplication = $null  # FIX: necesario para SaveAsText de formularios
    )

    $component = $null
    $tmp = $null
    $finalPath = $null

    try {
        $component = $VbProject.VBComponents.Item($ModuleName)
        $type = [int]$component.Type
        if ($type -ne 1 -and $type -ne 2 -and $type -ne 100 -and $type -ne 3) { return }
        $ext = Get-ComponentExtension -Component $component -ModuleName $ModuleName
        $folder = Get-ComponentFolder -Component $component -ModuleName $ModuleName
        if (-not $ext -or -not $folder) { return }

        $targetFolder = Join-Path -Path $ModulesPath -ChildPath $folder
        if (-not (Test-Path -Path $targetFolder)) {
            New-Item -Path $targetFolder -ItemType Directory -Force | Out-Null
        }

        $finalPath = Join-Path -Path $targetFolder -ChildPath ($ModuleName + $ext)

        # FIX: formularios usan SaveAsText para obtener UI + codigo completo
        # SaveAsText requiere el nombre del objeto Access SIN prefijo "Form_"
        if ($type -eq 3 -or $type -eq 100) {
            $formName = $ModuleName -replace '^Form_', ''
            $tmp = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("VBAManager_export_{0}.txt" -f [guid]::NewGuid().ToString("N"))

            if ($AccessApplication) {
                try {
                    # acForm = 2
                    $AccessApplication.SaveAsText(2, $formName, $tmp)
                } catch {
                    Write-Status -Message ("SaveAsText fallo para {0}: {1}" -f $formName, $_.Exception.Message) -Color Yellow
                    # FIX: limpiar $tmp anterior antes de reasignar
                    if ($tmp -and (Test-Path -Path $tmp)) { Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue }
                    $tmp = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("VBAManager_export_{0}.frm" -f [guid]::NewGuid().ToString("N"))
                    $component.Export($tmp)
                }
            } else {
                $component.Export($tmp)
            }
            Convert-AnsiToUtf8NoBom -InputPath $tmp -OutputPath $finalPath
        } else {
            $tmp = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("VBAManager_export_{0}{1}" -f @([guid]::NewGuid().ToString("N"), $ext))
            $component.Export($tmp)
            Convert-AnsiToUtf8NoBom -InputPath $tmp -OutputPath $finalPath
        }

        # Exportar tambien el codigo VBA como .cls para formularios (para diff y lectura rapida)
        if ($ModuleName -match "^Form_|^frm") {
            $clsFolder = Join-Path -Path $ModulesPath -ChildPath "forms"
            if (-not (Test-Path -Path $clsFolder)) {
                New-Item -Path $clsFolder -ItemType Directory -Force | Out-Null
            }
            $clsPath = Join-Path -Path $clsFolder -ChildPath ($ModuleName + ".cls")
            $codeModule = $component.CodeModule
            if ($codeModule -and $codeModule.CountOfLines -gt 0) {
                $codeLines = $codeModule.Lines(1, $codeModule.CountOfLines)
                Write-Utf8NoBom -Path $clsPath -Text $codeLines
            }
        }
    } finally {
        if ($tmp -and (Test-Path -Path $tmp)) { Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue }
        if ($component) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component) | Out-Null } catch {} }
    }
}

function Resolve-ImportFileForModule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [ValidateSet("Auto", "Form", "Code")][string]$ImportMode = "Auto"
    )

    $modulesPathText = [string]$ModulesPath
    $moduleNameText = [string]$ModuleName

    $subFolders = @("forms", "classes", "modules", "")
    switch ($ImportMode) {
        "Form" { $extensions = @(".form.txt", ".frm") }
        "Code" { $extensions = @(".cls", ".bas") }
        default { $extensions = @(".form.txt", ".frm", ".cls", ".bas") }
    }

    foreach ($folder in $subFolders) {
        $searchPath = if ($folder) { Join-Path -Path $modulesPathText -ChildPath $folder } else { $modulesPathText }
        if (-not (Test-Path -Path $searchPath)) { continue }

        foreach ($ext in $extensions) {
            $candidate = Join-Path -Path $searchPath -ChildPath ($moduleNameText + $ext)
            if (Test-Path -Path $candidate) { return $candidate }
        }
    }

    $any = Get-ChildItem -Path $modulesPathText -File -Recurse -Include "*.bas", "*.cls", "*.frm", "*.form.txt" -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -ieq $moduleNameText -or ($_.Name -replace '\.form\.txt$', '') -ieq $moduleNameText } |
        Where-Object {
            switch ($ImportMode) {
                "Form" { $_.Name -match '\.form\.txt$' -or $_.Extension -ieq '.frm' }
                "Code" { $_.Extension -ieq '.cls' -or $_.Extension -ieq '.bas' }
                default { $true }
            }
        } |
        Sort-Object -Property @{ Expression = {
            if ($ImportMode -eq "Code") {
                if ($_.Extension -eq '.cls') { 0 } elseif ($_.Extension -eq '.bas') { 1 } else { 9 }
            } else {
                if ($_.Name -match '\.form\.txt$') { 0 } elseif ($_.Extension -eq '.frm') { 1 } elseif ($_.Extension -eq '.cls') { 2 } else { 3 }
            }
        } } |
        Select-Object -First 1

    if ($any) { return $any.FullName }
    return $null
}

function Remove-ExistingComponent {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject,
        [Parameter(Mandatory = $true)][string]$ModuleName
    )

    $components = $VbProject.VBComponents
    for ($i = $components.Count; $i -ge 1; $i--) {
        $c = $components.Item($i)
        try {
            if ($c.Name -ieq $ModuleName) {
                $components.Remove($c)
                break
            }
        } finally {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($c) | Out-Null } catch {}
        }
    }
}

function Import-VbaModule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        $AccessApplication = $null,  # FIX: necesario para LoadFromText de formularios
        [ValidateSet("Auto", "Form", "Code")][string]$ImportMode = "Auto"
    )

    $src = Resolve-ImportFileForModule -ModulesPath $ModulesPath -ModuleName $ModuleName -ImportMode $ImportMode
    if (-not $src) { throw ("No se encontro archivo para el modulo '{0}' en {1}" -f $ModuleName, $ModulesPath) }

    $isFormTxt = ($src -match '\.form\.txt$')
    $ext = [System.IO.Path]::GetExtension($src)
    $tmpAnsi = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("VBAManager_import_{0}{1}" -f @([guid]::NewGuid().ToString("N"), $ext))
    $component = $null
    $codeModule = $null

    try {
        Convert-Utf8ToAnsiTempFile -InputPath $src -TempPath $tmpAnsi

        # FIX: formularios usan LoadFromText — nunca VBComponents.Import
        if ($isFormTxt) {
            if (-not $AccessApplication) { throw "Se necesita -AccessApplication para importar formularios (.form.txt)" }
            $formName = $ModuleName -replace '^Form_', ''
            try { $AccessApplication.DoCmd.SetWarnings($false) } catch {}
            # acForm = 2
            $AccessApplication.LoadFromText(2, $formName, $tmpAnsi)
            return
        }

        # FIX: .cls class modules require VBComponents.Import() to preserve metadata
        # (VERSION 1.0 CLASS, BEGIN...END, Attribute VB_Exposed, MultiUse, etc.)
        # CodeModule.AddFromFile() only adds code lines and strips class metadata.
        # .bas modules work fine with AddFromFile (no special metadata).
        $isClassModule = ($ext -eq ".cls")
        if ($isClassModule) {
            # Class modules: always use Import to preserve class metadata
            try {
                Remove-ExistingComponent -VbProject $VbProject -ModuleName $ModuleName
            } catch {}
            $VbProject.VBComponents.Import($tmpAnsi) | Out-Null
        } else {
            # Standard modules: DeleteLines + AddFromFile (safe for .bas)
            try {
                $component = $VbProject.VBComponents.Item($ModuleName)
                $codeModule = $component.CodeModule
                $count = $codeModule.CountOfLines
                if ($count -gt 0) { $codeModule.DeleteLines(1, $count) }
                $codeModule.AddFromFile($tmpAnsi)
            } catch {
                # El componente no existe aun — importar como nuevo (sin dialogo porque no hay nada que reemplazar)
                $VbProject.VBComponents.Import($tmpAnsi) | Out-Null
            }
        }

    } finally {
        if ($tmpAnsi -and (Test-Path -Path $tmpAnsi)) { Remove-Item -Path $tmpAnsi -Force -ErrorAction SilentlyContinue }
        foreach ($obj in @($codeModule, $component)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
    }
}

# ─── NUEVO: Delete-VbaModule ──────────────────────────────────────────
function Delete-VbaModule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject,
        [Parameter(Mandatory = $true)][string]$ModuleName,
        $AccessApplication = $null
    )

    $component = $null
    try {
        $component = $VbProject.VBComponents.Item($ModuleName)
    } catch {
        throw ("No se encontro el modulo '{0}' en el proyecto VBA." -f $ModuleName)
    }

    $type = [int]$component.Type

    try {
        # Formularios requieren DoCmd.DeleteObject (VBComponents.Remove no funciona para forms)
        if ($type -eq 3 -or $type -eq 100) {
            if (-not $AccessApplication) {
                throw "Se necesita AccessApplication para borrar formularios."
            }
            $formName = $ModuleName -replace '^Form_', ''
            try { $AccessApplication.DoCmd.SetWarnings($false) } catch {}
            # acForm = 2
            $AccessApplication.DoCmd.DeleteObject(2, $formName)
        } else {
            # Modulos y clases: VBComponents.Remove
            $VbProject.VBComponents.Remove($component)
        }
    } finally {
        if ($component) {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component) | Out-Null } catch {}
        }
    }
}

# ─── NUEVO: Rename-VbaModule ──────────────────────────────────────────
function Rename-VbaModule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject,
        [Parameter(Mandatory = $true)][string]$OldName,
        [Parameter(Mandatory = $true)][string]$NewName,
        $AccessApplication = $null
    )

    $component = $null
    try {
        $component = $VbProject.VBComponents.Item($OldName)
    } catch {
        throw ("No se encontro el modulo '{0}' en el proyecto VBA." -f $OldName)
    }

    $type = [int]$component.Type

    try {
        if ($type -eq 3 -or $type -eq 100) {
            # Formularios: usar DoCmd.Rename(NuevoNombre, acForm, NombreActual)
            if (-not $AccessApplication) {
                throw "Se necesita AccessApplication para renombrar formularios."
            }
            $oldFormName = $OldName -replace '^Form_', ''
            $newFormName = $NewName -replace '^Form_', ''
            try { $AccessApplication.DoCmd.SetWarnings($false) } catch {}
            # acForm = 2
            $AccessApplication.DoCmd.Rename($newFormName, 2, $oldFormName)
        } else {
            # Modulos y clases: cambiar la propiedad Name directamente
            $component.Name = $NewName
        }
    } finally {
        if ($component) {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component) | Out-Null } catch {}
        }
    }
}

# ─── NUEVO: Delete-SrcFilesForModule ──────────────────────────────────
function Delete-SrcFilesForModule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        [Parameter(Mandatory = $true)][string]$ModuleName
    )

    # Buscar todos los archivos que corresponden al modulo en cualquier subcarpeta
    $extensions = @(".form.txt", ".frm", ".cls", ".bas")
    $subFolders = @("forms", "classes", "modules", "")
    $deleted = 0

    foreach ($folder in $subFolders) {
        $searchPath = if ($folder) { Join-Path -Path $ModulesPath -ChildPath $folder } else { $ModulesPath }
        if (-not (Test-Path -Path $searchPath)) { continue }

        foreach ($ext in $extensions) {
            $candidate = Join-Path -Path $searchPath -ChildPath ($ModuleName + $ext)
            if (Test-Path -Path $candidate) {
                Remove-Item -LiteralPath $candidate -Force
                Write-Status -Message ("  Borrado: {0}" -f $candidate) -Color DarkYellow
                $deleted++
            }
        }
    }

    return $deleted
}

# ─── NUEVO: Rename-SrcFilesForModule ─────────────────────────────────
function Rename-SrcFilesForModule {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        [Parameter(Mandatory = $true)][string]$OldName,
        [Parameter(Mandatory = $true)][string]$NewName
    )

    $extensions = @(".form.txt", ".frm", ".cls", ".bas")
    $subFolders = @("forms", "classes", "modules", "")
    $renamed = 0

    foreach ($folder in $subFolders) {
        $searchPath = if ($folder) { Join-Path -Path $ModulesPath -ChildPath $folder } else { $ModulesPath }
        if (-not (Test-Path -Path $searchPath)) { continue }

        foreach ($ext in $extensions) {
            $oldFile = Join-Path -Path $searchPath -ChildPath ($OldName + $ext)
            if (Test-Path -Path $oldFile) {
                $newFile = Join-Path -Path $searchPath -ChildPath ($NewName + $ext)
                Rename-Item -LiteralPath $oldFile -NewName ([System.IO.Path]::GetFileName($newFile)) -Force
                Write-Status -Message ("  Renombrado: {0} -> {1}" -f ([System.IO.Path]::GetFileName($oldFile)), ([System.IO.Path]::GetFileName($newFile))) -Color DarkCyan
                $renamed++
            }
        }
    }

    return $renamed
}

function Fix-EncodingInSrc {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        [string[]]$ModuleName
    )

    $targets = @()
    if ($ModuleName -and $ModuleName.Count -gt 0) {
        foreach ($m in $ModuleName) {
            $f = Resolve-ImportFileForModule -ModulesPath $ModulesPath -ModuleName $m
            if ($f) { $targets += $f }
        }
    } else {
        $targets = @(Get-ChildItem -Path $ModulesPath -Recurse -File -Include "*.bas", "*.cls", "*.frm", "*.form.txt" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
    }

    $fixed = 0
    foreach ($p in $targets) {
        $utf8 = [System.Text.Encoding]::UTF8
        $text = [System.IO.File]::ReadAllText($p, $utf8)
        $info = Get-FileEncodingInfo -Path $p
        if ($info.HasUtf8Bom) {
            Write-Utf8NoBom -Path $p -Text $text
            $fixed++
        }
    }
    return $fixed
}

function Export-DataStructure {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$DatabasePath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$Password = ""
    )

    $dbEngine = $null
    $database = $null

    $typeMap = @{
        1 = "Boolean"; 2 = "Byte"; 3 = "Integer"; 4 = "Long"; 5 = "Currency"
        6 = "Single"; 7 = "Double"; 8 = "Date/Time"; 9 = "Binary"; 10 = "Text"
        11 = "OLE"; 12 = "Memo"; 15 = "GUID"; 16 = "BigInt"
        17 = "VarBinary"; 18 = "Char"; 19 = "Numeric"; 20 = "Decimal"
    }

    try {
        $dbEngine = New-DaoDbEngine
        if (-not $dbEngine) { throw "No se pudo crear DAO.DBEngine" }

        $connect = if (-not [string]::IsNullOrEmpty($Password)) { ";PWD=$Password" } else { "" }
        $database = $dbEngine.OpenDatabase($DatabasePath, $false, $true, $connect)

        $sb = [System.Text.StringBuilder]::new()
        $dbName = [System.IO.Path]::GetFileNameWithoutExtension($DatabasePath)
        [void]$sb.AppendLine("# ERD - $dbName")
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Generado: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
        [void]$sb.AppendLine("")

        # ── Paso 1: recopilar metadatos de todas las tablas ──────────────
        # DAO Attributes: dbAttachedTable = 0x40000000, dbAttachedODBC = 0x20000000
        $dbAttachedTable = 1073741824
        $dbAttachedODBC  = 536870912

        $tableDefs = $database.TableDefs
        $tableInfos = @()               # lista de objetos con metadatos
        $linkedEntries = @()             # solo las vinculadas

        for ($i = 0; $i -lt $tableDefs.Count; $i++) {
            $td = $tableDefs[$i]
            try {
                if ($td.Name -match "^MSys" -or $td.Name -match "^~") { continue }

                $attrs      = 0
                $tdConnect  = ""
                $srcTable   = ""
                try { $attrs     = [long]$td.Attributes }  catch {}
                try { $tdConnect = [string]$td.Connect }   catch {}
                try { $srcTable  = [string]$td.SourceTableName } catch {}

                $isLinkedAccess = ($attrs -band $dbAttachedTable) -ne 0
                $isLinkedODBC   = ($attrs -band $dbAttachedODBC)  -ne 0
                $isLinked       = $isLinkedAccess -or $isLinkedODBC -or (-not [string]::IsNullOrEmpty($tdConnect))

                # Clasificar tipo de vinculacion
                $linkType   = ""
                $linkTarget = ""
                if ($isLinked -and -not [string]::IsNullOrEmpty($tdConnect)) {
                    if ($isLinkedODBC -or $tdConnect -match "^ODBC;") {
                        $linkType = "ODBC"
                        # Intentar extraer DSN o DRIVER
                        if ($tdConnect -match "DSN=([^;]+)") {
                            $linkTarget = "DSN=$($Matches[1])"
                        } elseif ($tdConnect -match "DRIVER=\{?([^;}]+)") {
                            $linkTarget = "Driver=$($Matches[1])"
                        }
                        # Intentar extraer SERVER/DATABASE para dar mas contexto
                        $serverPart = ""
                        $dbPart = ""
                        if ($tdConnect -match "SERVER=([^;]+)") { $serverPart = $Matches[1] }
                        if ($tdConnect -match "DATABASE=([^;]+)") { $dbPart = $Matches[1] }
                        if ($serverPart -and $dbPart) {
                            $linkTarget += " ($serverPart/$dbPart)"
                        } elseif ($dbPart) {
                            $linkTarget += " ($dbPart)"
                        }
                    } elseif ($tdConnect -match ";DATABASE=(.+)$") {
                        $linkType = "Access"
                        $linkTarget = $Matches[1].Trim()
                    } elseif ($tdConnect -match "^Excel") {
                        $linkType = "Excel"
                        if ($tdConnect -match "DATABASE=(.+)$") { $linkTarget = $Matches[1].Trim() }
                    } elseif ($tdConnect -match "^SharePoint" -or $tdConnect -match "WSS;") {
                        $linkType = "SharePoint"
                        if ($tdConnect -match "DATABASE=(.+)$") { $linkTarget = $Matches[1].Trim() }
                        elseif ($tdConnect -match "LIST=([^;]+)") { $linkTarget = "Lista: $($Matches[1])" }
                    } elseif ($tdConnect -match "^Text;") {
                        $linkType = "Text/CSV"
                        if ($tdConnect -match "DATABASE=(.+)$") { $linkTarget = $Matches[1].Trim() }
                    } elseif ($tdConnect -match "^HTML") {
                        $linkType = "HTML"
                        if ($tdConnect -match "DATABASE=(.+)$") { $linkTarget = $Matches[1].Trim() }
                    } else {
                        $linkType = "Otro"
                        $linkTarget = $tdConnect
                    }
                } elseif ($isLinked) {
                    $linkType = "Desconocido"
                    $linkTarget = "(Connect vacio, Attributes=$attrs)"
                }

                $info = [pscustomobject]@{
                    Name            = $td.Name
                    IsLinked        = $isLinked
                    LinkType        = $linkType
                    LinkTarget      = $linkTarget
                    SourceTableName = $srcTable
                    ConnectString   = $tdConnect
                }
                $tableInfos += $info
                if ($isLinked) { $linkedEntries += $info }
            } catch {}
        }

        $tableInfos = $tableInfos | Sort-Object -Property Name
        $localTables  = @($tableInfos | Where-Object { -not $_.IsLinked })
        $linkedTables = @($tableInfos | Where-Object { $_.IsLinked })

        [void]$sb.AppendLine("## Tablas ($($tableInfos.Count) total: $($localTables.Count) locales, $($linkedTables.Count) vinculadas)")
        [void]$sb.AppendLine("")

        # ── Paso 2: detalle de cada tabla ─────────────────────────────────
        foreach ($tInfo in $tableInfos) {
            try {
                $td = $database.TableDefs[$tInfo.Name]
                $header = $tInfo.Name
                if ($tInfo.IsLinked) {
                    $linkedLabel = $tInfo.LinkType
                    if ($tInfo.SourceTableName -and $tInfo.SourceTableName -ne $tInfo.Name) {
                        $linkedLabel += ", origen: ``$($tInfo.SourceTableName)``"
                    }
                    $header = "$($tInfo.Name) _(vinculada: $linkedLabel)_"
                }
                [void]$sb.AppendLine("### $header")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("| Campo | Tipo | Tamaño | Requerido | PK |")
                [void]$sb.AppendLine("|---|---|---|---|---|")

                $pkFields = @()
                try {
                    for ($i = 0; $i -lt $td.Indexes.Count; $i++) {
                        $idx = $td.Indexes[$i]
                        if ($idx.Primary) {
                            for ($j = 0; $j -lt $idx.Fields.Count; $j++) {
                                $pkFields += $idx.Fields[$j].Name
                            }
                        }
                    }
                } catch {}

                for ($i = 0; $i -lt $td.Fields.Count; $i++) {
                    try {
                        $field = $td.Fields[$i]
                        $typeCode = [int]$field.Type
                        $typeName = if ($typeMap.ContainsKey($typeCode)) { $typeMap[$typeCode] } else { "Tipo$typeCode" }
                        $size = if ($field.Size -gt 0) { $field.Size } else { "-" }
                        $required = if ($field.Required) { "Si" } else { "No" }
                        $isPk = if ($pkFields -contains $field.Name) { "PK" } else { "" }
                        [void]$sb.AppendLine("| $($field.Name) | $typeName | $size | $required | $isPk |")
                    } catch {}
                }
                [void]$sb.AppendLine("")
            } catch {
                [void]$sb.AppendLine("_Error leyendo tabla: $($tInfo.Name) - $($_.Exception.Message)_")
                [void]$sb.AppendLine("")
            }
        }

        # ── Paso 3: relaciones ────────────────────────────────────────────
        try {
            $relations = $database.Relations
            if ($relations.Count -gt 0) {
                [void]$sb.AppendLine("## Relaciones")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("| Nombre | Tabla origen | Campo origen | Tabla destino | Campo destino |")
                [void]$sb.AppendLine("|---|---|---|---|---|")

                for ($i = 0; $i -lt $relations.Count; $i++) {
                    try {
                        $rel = $relations[$i]
                        $originField = ""
                        $foreignField = ""
                        if ($rel.Fields.Count -gt 0) {
                            $rf = $rel.Fields[0]
                            $originField = $rf.Name
                            $foreignField = $rf.ForeignName
                        }
                        [void]$sb.AppendLine("| $($rel.Name) | $($rel.Table) | $originField | $($rel.ForeignTable) | $foreignField |")
                    } catch {}
                }
                [void]$sb.AppendLine("")
            }
        } catch {}

        # ── Paso 4: resumen de tablas vinculadas ──────────────────────────
        if ($linkedEntries.Count -gt 0) {
            [void]$sb.AppendLine("## Tablas vinculadas")
            [void]$sb.AppendLine("")

            # Agrupar por tipo de vinculacion
            $byType = $linkedEntries | Group-Object -Property LinkType | Sort-Object -Property Name
            foreach ($group in $byType) {
                [void]$sb.AppendLine("### Vinculadas por $($group.Name)")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("| Tabla local | Tabla origen | Destino |")
                [void]$sb.AppendLine("|---|---|---|")

                foreach ($entry in $group.Group | Sort-Object -Property Name) {
                    $srcName = if ($entry.SourceTableName -and $entry.SourceTableName -ne $entry.Name) { $entry.SourceTableName } else { "—" }
                    $target  = if ($entry.LinkTarget) { "``$($entry.LinkTarget)``" } else { "—" }
                    [void]$sb.AppendLine("| $($entry.Name) | $srcName | $target |")
                }
                [void]$sb.AppendLine("")
            }

            # Comprobar alcanzabilidad de origenes basados en fichero
            $fileBasedTypes = @("Access", "Excel", "Text/CSV", "HTML")
            $fileTargets = @{}
            foreach ($entry in $linkedEntries) {
                if ($entry.LinkType -in $fileBasedTypes -and $entry.LinkTarget) {
                    if (-not $fileTargets.ContainsKey($entry.LinkTarget)) {
                        $fileTargets[$entry.LinkTarget] = [System.Collections.Generic.List[string]]::new()
                    }
                    $fileTargets[$entry.LinkTarget].Add($entry.Name)
                }
            }

            $unreachable = @($fileTargets.Keys | Where-Object { -not (Test-Path -Path $_) })
            if ($unreachable.Count -gt 0) {
                [void]$sb.AppendLine("### Origenes no alcanzados")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("Los siguientes origenes de datos vinculados no estaban disponibles al generar este ERD.")
                [void]$sb.AppendLine("Sus tablas aparecen en el listado pero su estructura no pudo verificarse.")
                [void]$sb.AppendLine("")
                foreach ($path in $unreachable | Sort-Object) {
                    $affectedTables = $fileTargets[$path] -join ", "
                    [void]$sb.AppendLine("- ``$path`` — tablas: $affectedTables")
                }
                [void]$sb.AppendLine("")
            }

            # Origenes ODBC (no se puede comprobar alcanzabilidad por fichero)
            $odbcEntries = @($linkedEntries | Where-Object { $_.LinkType -eq "ODBC" })
            if ($odbcEntries.Count -gt 0) {
                $odbcTargets = @($odbcEntries | ForEach-Object { $_.LinkTarget } | Sort-Object -Unique)
                [void]$sb.AppendLine("### Conexiones ODBC")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("La verificacion de alcanzabilidad de conexiones ODBC no es posible desde DAO.")
                [void]$sb.AppendLine("Destinos detectados:")
                [void]$sb.AppendLine("")
                foreach ($target in $odbcTargets) {
                    $odbcTablesForTarget = @($odbcEntries | Where-Object { $_.LinkTarget -eq $target } | ForEach-Object { $_.Name }) -join ", "
                    [void]$sb.AppendLine("- ``$target`` — tablas: $odbcTablesForTarget")
                }
                [void]$sb.AppendLine("")
            }
        }

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), $utf8NoBom)

    } finally {
        if ($database) {
            try { $database.Close() } catch {}
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($database) | Out-Null } catch {}
        }
        if ($dbEngine) {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dbEngine) | Out-Null } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Fix-EncodingInAccess {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject,
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        [string[]]$ModuleName,
        $AccessApplication = $null
    )

    $components = $VbProject.VBComponents
    $names = @()

    if ($ModuleName -and $ModuleName.Count -gt 0) {
        $names = @($ModuleName | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } else {
        for ($i = 1; $i -le $components.Count; $i++) {
            $c = $components.Item($i)
            try {
                $type = [int]$c.Type
                if ($type -ne 1 -and $type -ne 2 -and $type -ne 100 -and $type -ne 3) { continue }
                $ext = Get-ComponentExtension -Component $c -ModuleName $c.Name
                if ($ext) { $names += $c.Name }
            } finally {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($c) | Out-Null } catch {}
            }
        }
    }

    $fixed = 0
    foreach ($n in $names | Sort-Object -Unique) {
        try {
            Export-VbaModule -VbProject $VbProject -ModuleName $n -ModulesPath $ModulesPath -AccessApplication $AccessApplication
            Import-VbaModule -VbProject $VbProject -ModuleName $n -ModulesPath $ModulesPath -AccessApplication $AccessApplication
            $fixed++
        } catch {
            Write-Status -Message ("ERROR en modulo '{0}': {1}" -f $n, $_.Exception.Message) -Color Red
        }
    }
    return $fixed
}

# ═══════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════

$session = $null

try {
    $DestinationRoot = Resolve-DestinationRoot -DestinationRoot $DestinationRoot

    if ($Action -notin @("Generate-ERD", "List")) {
        $AccessPath = Resolve-AccessPath -AccessPath $AccessPath
        $ModulesPath = Resolve-ModulesPath -DestinationRoot $DestinationRoot -AccessPath $AccessPath -Action $Action

        Write-Status -Message ("Accion: {0}" -f $Action) -Color Yellow
        Write-Status -Message ("Base de datos: {0}" -f $AccessPath) -Color Yellow
        Write-Status -Message ("Carpeta: {0}" -f $ModulesPath) -Color Yellow
    } elseif ($Action -eq "List") {
        $AccessPath = Resolve-AccessPath -AccessPath $AccessPath
        Write-Status -Message ("Accion: {0}" -f $Action) -Color Yellow
        Write-Status -Message ("Base de datos: {0}" -f $AccessPath) -Color Yellow
    } elseif ($Action -eq "Generate-ERD") {
        # Generate-ERD resuelve sus propias rutas en el bloque de accion
        # $AccessPath se pasa tal cual (puede estar vacio)
        Write-Status -Message ("Accion: {0}" -f $Action) -Color Yellow
    } else {
        Write-Status -Message ("Accion: {0}" -f $Action) -Color Yellow
    }

    # FIX: Normalize ModuleName - handle comma-separated input from handler join
    # When handler passes "Mod1,Mod2,Mod3" as single string, it becomes String[] with one element
    $inputModules = $ModuleName
    if ($inputModules.Count -eq 1 -and $inputModules[0] -is [string] -and $inputModules[0] -match ',') {
        $inputModules = $inputModules[0] -split ','
    }
    $normalizedModules = @($inputModules | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($Action -eq "Export") {
        $session = Open-AccessDatabase -AccessPath $AccessPath -Password $Password
        $vbProject = $session.VbProject
        $components = $vbProject.VBComponents

        $targets = @()
        if ($normalizedModules.Count -gt 0) {
            $targets = $normalizedModules
        } else {
            for ($i = 1; $i -le $components.Count; $i++) {
                $c = $components.Item($i)
                try {
                    $ext = Get-ComponentExtension -Component $c -ModuleName $c.Name
                    if ($ext) { $targets += $c.Name }
                } finally {
                    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($c) | Out-Null } catch {}
                }
            }
            $targets = $targets | Sort-Object -Unique
        }

        $total = $targets.Count
        $idx = 0
        foreach ($name in $targets) {
            $idx++
            Write-Status -Message ("[{0}/{1}] Exportando: {2}" -f $idx, $total, $name) -Color Cyan
            # FIX: pasar AccessApplication para que SaveAsText funcione en formularios
            Export-VbaModule -VbProject $vbProject -ModuleName $name -ModulesPath $ModulesPath -AccessApplication $session.AccessApplication
        }
        Write-Status -Message ("OK Export completado ({0})" -f $total) -Color Green

    } elseif ($Action -eq "Import") {
        $session = Open-AccessDatabase -AccessPath $AccessPath -Password $Password
        $vbProject = $session.VbProject

        $targets = @()
        if ($normalizedModules.Count -gt 0) {
            $targets = $normalizedModules
        } else {
            # FIX: incluir *.form.txt y extraer nombre correctamente
            $targets = @(Get-ChildItem -Path $ModulesPath -File -Recurse `
                -Include "*.bas", "*.cls", "*.frm", "*.form.txt" -ErrorAction SilentlyContinue |
                ForEach-Object {
                    if ($_.Name -match '\.form\.txt$') { $_.Name -replace '\.form\.txt$', '' }
                    else { $_.BaseName }
                } | Sort-Object -Unique)
        }

        $total = $targets.Count
        $idx = 0
        foreach ($name in $targets) {
            $idx++
            Write-Status -Message ("[{0}/{1}] Importando: {2}" -f $idx, $total, $name) -Color Cyan
            # FIX: pasar AccessApplication para LoadFromText en formularios
            Import-VbaModule -VbProject $vbProject -ModuleName $name -ModulesPath $ModulesPath -AccessApplication $session.AccessApplication -ImportMode $ImportMode
        }
        Write-Status -Message ("OK Import completado ({0})" -f $total) -Color Green

    # ─── NUEVO: Delete ─────────────────────────────────────────────────
    } elseif ($Action -eq "Delete") {
        if ($normalizedModules.Count -eq 0) {
            throw "Se necesita al menos un nombre de modulo para Delete."
        }

        $session = Open-AccessDatabase -AccessPath $AccessPath -Password $Password

        $total = $normalizedModules.Count
        $idx = 0
        foreach ($name in $normalizedModules) {
            $idx++
            Write-Status -Message ("[{0}/{1}] Borrando: {2}" -f $idx, $total, $name) -Color Magenta
            try {
                Delete-VbaModule -VbProject $session.VbProject -ModuleName $name -AccessApplication $session.AccessApplication
                Write-Status -Message ("  OK borrado de BD: {0}" -f $name) -Color Green

                if ($DeleteFromSrc) {
                    $deletedFiles = Delete-SrcFilesForModule -ModulesPath $ModulesPath -ModuleName $name
                    if ($deletedFiles -eq 0) {
                        Write-Status -Message ("  No se encontraron archivos en src/ para: {0}" -f $name) -Color DarkYellow
                    }
                }
            } catch {
                Write-Status -Message ("  ERROR borrando '{0}': {1}" -f $name, $_.Exception.Message) -Color Red
            }
        }
        Write-Status -Message ("OK Delete completado ({0})" -f $total) -Color Green

    # ─── NUEVO: Rename ─────────────────────────────────────────────────
    } elseif ($Action -eq "Rename") {
        if ($normalizedModules.Count -ne 1) {
            throw "Rename requiere exactamente un modulo en -ModuleName (el nombre actual)."
        }
        if ([string]::IsNullOrWhiteSpace($NewModuleName)) {
            throw "Rename requiere -NewModuleName con el nuevo nombre."
        }

        $oldName = $normalizedModules[0]
        $session = Open-AccessDatabase -AccessPath $AccessPath -Password $Password

        Write-Status -Message ("Renombrando: {0} -> {1}" -f $oldName, $NewModuleName) -Color Cyan
        Rename-VbaModule -VbProject $session.VbProject -OldName $oldName -NewName $NewModuleName -AccessApplication $session.AccessApplication
        Write-Status -Message ("  OK renombrado en BD") -Color Green

        # Renombrar archivos en src/
        $renamedFiles = Rename-SrcFilesForModule -ModulesPath $ModulesPath -OldName $oldName -NewName $NewModuleName
        if ($renamedFiles -eq 0) {
            Write-Status -Message ("  No se encontraron archivos en src/ para renombrar.") -Color DarkYellow
            Write-Status -Message ("  Ejecuta 'export {0}' para generar los archivos con el nuevo nombre." -f $NewModuleName) -Color Yellow
        }

        Write-Status -Message ("OK Rename completado") -Color Green

    # ─── NUEVO: List ───────────────────────────────────────────────────
    } elseif ($Action -eq "List") {
        $session = Open-AccessDatabase -AccessPath $AccessPath -Password $Password
        $vbProject = $session.VbProject
        $components = $vbProject.VBComponents

        $modules = @()
        for ($i = 1; $i -le $components.Count; $i++) {
            $c = $components.Item($i)
            try {
                $type = [int]$c.Type
                $typeName = Get-ComponentTypeName -Type $type
                $lineCount = 0
                try {
                    if ($c.CodeModule) { $lineCount = $c.CodeModule.CountOfLines }
                } catch {}
                $modules += [pscustomobject]@{
                    Name  = $c.Name
                    Type  = $typeName
                    Lines = $lineCount
                }
            } finally {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($c) | Out-Null } catch {}
            }
        }

        $modules = $modules | Sort-Object -Property Type, Name

        Write-Host ""
        Write-Host ("  {0,-40} {1,-10} {2,6}" -f "Nombre", "Tipo", "Lineas")
        Write-Host ("  {0,-40} {1,-10} {2,6}" -f ("─" * 40), ("─" * 10), ("─" * 6))
        foreach ($m in $modules) {
            $color = switch ($m.Type) {
                "Form"   { [ConsoleColor]::Cyan }
                "Class"  { [ConsoleColor]::Yellow }
                "Module" { [ConsoleColor]::Green }
                default  { [ConsoleColor]::Gray }
            }
            $old = $Host.UI.RawUI.ForegroundColor
            $Host.UI.RawUI.ForegroundColor = $color
            Write-Host ("  {0,-40} {1,-10} {2,6}" -f $m.Name, $m.Type, $m.Lines)
            $Host.UI.RawUI.ForegroundColor = $old
        }
        Write-Host ""
        Write-Status -Message ("Total: {0} modulo(s)" -f $modules.Count) -Color Green

    } elseif ($Action -eq "Generate-ERD") {
        # ── Resolver qué BDs escanear ─────────────────────────────────────
        # --backend: genera ERD del backend (comportamiento clasico)
        # --access:  genera ERD del frontend (util para ver tablas vinculadas)
        # ambos:     genera ERD de los dos
        # ninguno:   auto-detecta backend *_Datos.accdb; si no hay, intenta frontend

        $dbsToScan = @()   # lista de [pscustomobject]@{ Path; Label }

        $hasExplicitBackend = -not [string]::IsNullOrWhiteSpace($BackendPath)
        $hasExplicitAccess  = -not [string]::IsNullOrWhiteSpace($AccessPath)

        if ($hasExplicitBackend) {
            $BackendPath = (Resolve-Path -Path $BackendPath).Path
            $dbsToScan += [pscustomobject]@{ Path = $BackendPath; Label = "Backend" }
        }

        if ($hasExplicitAccess) {
            $AccessPath = (Resolve-Path -Path $AccessPath).Path
            $dbsToScan += [pscustomobject]@{ Path = $AccessPath; Label = "Frontend" }
        }

        # Si no se paso ninguno, auto-detectar
        if ($dbsToScan.Count -eq 0) {
            $candidates = Get-ChildItem -Path (Get-Location) -File -Filter "*_Datos.accdb" -ErrorAction SilentlyContinue
            if (-not $candidates) {
                $candidates = Get-ChildItem -Path (Get-Location) -File -Filter "*_Datos.mdb" -ErrorAction SilentlyContinue
            }

            if ($candidates) {
                if ($candidates.Count -gt 1) {
                    Write-Status -Message "ADVERTENCIA: Multiples backends encontrados, usando el primero: $($candidates[0].Name)" -Color Yellow
                }
                $dbsToScan += [pscustomobject]@{ Path = $candidates[0].FullName; Label = "Backend" }
            } else {
                # Sin backend: intentar el frontend (la BD principal)
                try {
                    $frontendPath = (Resolve-AccessPath -AccessPath "").Trim()
                    if ($frontendPath) {
                        $dbsToScan += [pscustomobject]@{ Path = $frontendPath; Label = "Frontend" }
                    }
                } catch {}
            }

            if ($dbsToScan.Count -eq 0) {
                throw "No se especifico -BackendPath ni -AccessPath y no se encontro ningun archivo .accdb/.mdb en el directorio actual."
            }
        }

        # ── Resolver carpeta de salida ────────────────────────────────────
        if ([string]::IsNullOrWhiteSpace($ErdPath)) {
            $parent = Split-Path -Parent $DestinationRoot
            $ErdPath = Join-Path -Path $parent -ChildPath "ERD"
        }

        if (-not (Test-Path -Path $ErdPath)) {
            New-Item -ItemType Directory -Force -Path $ErdPath | Out-Null
        }
        $ErdPath = (Resolve-Path -Path $ErdPath).Path
        Write-Status -Message ("ERD Folder: {0}" -f $ErdPath) -Color Yellow

        # ── Generar ERD para cada BD ──────────────────────────────────────
        foreach ($dbInfo in $dbsToScan) {
            $dbPath = $dbInfo.Path
            $dbLabel = $dbInfo.Label
            $dbBaseName = [System.IO.Path]::GetFileNameWithoutExtension($dbPath)
            $mdFile = Join-Path -Path $ErdPath -ChildPath ($dbBaseName + ".md")

            Write-Status -Message ("{0}: {1}" -f $dbLabel, $dbPath) -Color Yellow
            Export-DataStructure -DatabasePath $dbPath -OutputPath $mdFile -Password $Password
            Write-Status -Message ("  -> {0}" -f $mdFile) -Color Green
        }

        Write-Status -Message ("OK ERD generado ({0} fichero(s))" -f $dbsToScan.Count) -Color Green

    } else {
        # Fix-Encoding
        $fixedSrc = 0
        $fixedAccess = 0

        if ($Location -eq "Src" -or $Location -eq "Both") {
            $fixedSrc = Fix-EncodingInSrc -ModulesPath $ModulesPath -ModuleName $normalizedModules
            Write-Status -Message ("Fix-Encoding (Src): {0}" -f $fixedSrc) -Color Yellow
        }

        if ($Location -eq "Access" -or $Location -eq "Both") {
            $session = Open-AccessDatabase -AccessPath $AccessPath -Password $Password
            $fixedAccess = Fix-EncodingInAccess -VbProject $session.VbProject -ModulesPath $ModulesPath -ModuleName $normalizedModules -AccessApplication $session.AccessApplication
            Write-Status -Message ("Fix-Encoding (Access): {0}" -f $fixedAccess) -Color Yellow
        }

        Write-Status -Message ("OK Fix-Encoding completado") -Color Green
    }
} finally {
    if ($session) {
        try { Close-AccessDatabase -Session $session -AccessPath $AccessPath -Password $Password } catch {}
    }
}
