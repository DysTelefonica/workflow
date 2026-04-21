[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("Export", "Import", "Fix-Encoding", "Generate-ERD", "Sandbox")]
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

    [Parameter()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "Password", Justification = "Requerido por especificacion del proyecto.")]
    [string]$Password = "",

    [Parameter()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "BackendPassword", Justification = "Requerido por especificacion del proyecto.")]
    [string]$BackendPassword,

    [Parameter()]
    [switch]$KeepSidecars
)

$ErrorActionPreference = "Stop"
$script:AnsiCodePage = 1252

function Write-Status {
    Param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    Write-Host $Message -ForegroundColor $Color
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

function Invoke-WithDaoDatabase {
    # Abre una BD via DAO, ejecuta el scriptblock, y limpia COM/GC.
    # El scriptblock recibe $db como argumento.
    # Devuelve lo que devuelva el scriptblock, o $DefaultOnError si falla.
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password,
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        $DefaultOnError = $null
    )

    $dbEngine = $null
    $db = $null
    try {
        $dbEngine = New-DaoDbEngine
        if (-not $dbEngine) { return $DefaultOnError }
        $connect = if ($Password) { ";PWD=$Password" } else { "" }
        try { $db = $dbEngine.OpenDatabase($AccessPath, $false, $false, $connect) } catch { return $DefaultOnError }
        return (& $Action $db)
    } catch {
        return $DefaultOnError
    } finally {
        if ($db) { try { $db.Close() } catch {} }
        foreach ($obj in @($db, $dbEngine)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-AllowBypassKeyState {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
        [string]$Password
    )
    Invoke-WithDaoDatabase -AccessPath $AccessPath -Password $Password -Action {
        param($db)
        try {
            $prop = $db.Properties("AllowBypassKey")
            return [pscustomobject]@{ Existed = $true; Value = [bool]$prop.Value }
        } catch {
            return [pscustomobject]@{ Existed = $false; Value = $null }
        }
    }
}

function Enable-AllowBypassKey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Justification = "Requerido por especificacion del proyecto.")]
        [string]$Password
    )
    $result = Invoke-WithDaoDatabase -AccessPath $AccessPath -Password $Password -DefaultOnError $false -Action {
        param($db)
        try {
            $prop = $db.Properties("AllowBypassKey")
            $prop.Value = $true
        } catch {
            $newProp = $db.CreateProperty("AllowBypassKey", 1, $true)
            $db.Properties.Append($newProp)
        }
        return $true
    }
    return $result
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
    Invoke-WithDaoDatabase -AccessPath $AccessPath -Password $Password -Action {
        param($db)
        if ($OriginalState.Existed) {
            $prop = $db.Properties("AllowBypassKey")
            $prop.Value = [bool]$OriginalState.Value
        } else {
            try { $db.Properties.Delete("AllowBypassKey") } catch {}
        }
    } | Out-Null
}

function Disable-StartupFeatures {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [string]$Password
    )
    Invoke-WithDaoDatabase -AccessPath $AccessPath -Password $Password -Action {
        param($db)
        $restoreInfo = [ordered]@{
            RenamedAutoExec     = $false
            OriginalStartupForm = $null
            HasStartupForm      = $false
        }
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
    Invoke-WithDaoDatabase -AccessPath $AccessPath -Password $Password -Action {
        param($db)
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
    } | Out-Null
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
        [Parameter(Mandatory = $true)][ValidateSet("Export", "Import", "Fix-Encoding", "Generate-ERD", "Sandbox")][string]$Action
    )
    if (-not (Test-Path -Path $DestinationRoot)) {
        if ($Action -eq "Export" -or $Action -eq "Fix-Encoding") {
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

    $ansi = [System.Text.Encoding]::GetEncoding($script:AnsiCodePage)
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
    $ansi = [System.Text.Encoding]::GetEncoding($script:AnsiCodePage)
    $text = [System.IO.File]::ReadAllText($InputPath, $utf8)
    [System.IO.File]::WriteAllText($TempPath, $text, $ansi)
}

function Strip-VbaMetadataHeader {
    # Elimina del archivo las lineas de metadatos VBE que preceden al codigo real:
    #   VERSION 1.0 CLASS, bloque BEGIN/END, Attribute VB_*
    # Necesario antes de AddFromFile, que no parsea metadatos y los inyecta como codigo.
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$FilePath
    )

    $ansi = [System.Text.Encoding]::GetEncoding($script:AnsiCodePage)
    $lines = [System.IO.File]::ReadAllLines($FilePath, $ansi)
    $startIdx = 0
    $inBeginBlock = $false
    $foundMeta = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].TrimStart()

        if ($inBeginBlock) {
            if ($line -eq 'END') { $inBeginBlock = $false }
            continue
        }

        if ($line -match '^VERSION\s+') { $foundMeta = $true; continue }
        if ($line -eq 'BEGIN') { $foundMeta = $true; $inBeginBlock = $true; continue }
        if ($line -match '^Attribute\s+VB_') { $foundMeta = $true; continue }
        if ($foundMeta -and $line -eq '') { continue }  # solo saltar vacias entre metadatos

        $startIdx = $i
        break
    }

    if ($startIdx -gt 0) {
        $codeLines = $lines[$startIdx..($lines.Count - 1)]
        [System.IO.File]::WriteAllLines($FilePath, $codeLines, $ansi)
    }
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

function Close-TargetAccessDbIfOpen {
    # Cierra SOLO la instancia COM de Access que tiene abierta la BD indicada,
    # iterando el ROT completo para no afectar otras instancias de Access en ejecucion.
    # Toda la interaccion COM se hace en C# para evitar el problema de __ComObject opaco.
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)][string]$AccessPath
    )

    $resolved = $null
    $rp = Resolve-Path -Path $AccessPath -ErrorAction SilentlyContinue
    if ($rp) { $resolved = $rp.Path }
    # Fallback: si Resolve-Path falla (OneDrive, rutas largas), usar el path raw
    if (-not $resolved) {
        if (Test-Path -LiteralPath $AccessPath) { $resolved = $AccessPath }
        else {
            Write-Status -Message ("Close-TargetAccessDbIfOpen: no se pudo resolver la ruta: {0}" -f $AccessPath) -Color DarkYellow
            return
        }
    }

    # Registrar tipos solo una vez por sesion de PowerShell
    if (-not ([System.Management.Automation.PSTypeName]"RotManager").Type) {
        Add-Type -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class RotCloseResult {
    public bool Success;
    public string Error;
    public int ClosedCount;
}

public class RotManager {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

    public static RotCloseResult CloseDatabaseIfOpen(string dbPath) {
        var result = new RotCloseResult { Success = true };
        IRunningObjectTable rot = null;
        IEnumMoniker enumMk = null;
        IBindCtx bindCtx = null;

        try {
            int hr = GetRunningObjectTable(0, out rot);
            if (hr != 0 || rot == null) { result.Error = "No se pudo obtener el ROT"; return result; }

            hr = CreateBindCtx(0, out bindCtx);
            if (hr != 0 || bindCtx == null) { result.Error = "No se pudo crear BindCtx"; return result; }

            rot.EnumRunning(out enumMk);
            if (enumMk == null) { result.Error = "EnumRunning devolvio null"; return result; }

            enumMk.Reset();
            var monikers = new IMoniker[1];

            while (enumMk.Next(1, monikers, IntPtr.Zero) == 0) {
                if (monikers[0] == null) continue;
                object comObj = null;
                try {
                    string displayName = null;
                    try { monikers[0].GetDisplayName(bindCtx, null, out displayName); } catch { continue; }
                    if (string.IsNullOrEmpty(displayName) || !displayName.Contains("Access.Application")) continue;

                    try { rot.GetObject(monikers[0], out comObj); } catch { continue; }
                    if (comObj == null) continue;

                    // Usar reflection (late-binding) — funciona sobre __ComObject sin interop assembly
                    object db = null;
                    string openDbName = null;
                    try {
                        db = comObj.GetType().InvokeMember("CurrentDb",
                            BindingFlags.InvokeMethod, null, comObj, null);
                        if (db != null) {
                            openDbName = (string)db.GetType().InvokeMember("Name",
                                BindingFlags.GetProperty, null, db, null);
                        }
                    } catch {
                        // No tiene BD abierta o instancia corrupta — saltar
                    } finally {
                        if (db != null) try { Marshal.ReleaseComObject(db); } catch {}
                    }

                    if (!string.IsNullOrEmpty(openDbName) &&
                        string.Equals(openDbName, dbPath, StringComparison.OrdinalIgnoreCase)) {
                        try {
                            comObj.GetType().InvokeMember("CloseCurrentDatabase",
                                BindingFlags.InvokeMethod, null, comObj, null);
                            result.ClosedCount++;
                        } catch {}
                    }
                } catch {
                    // Este moniker no sirve — continuar
                } finally {
                    if (comObj != null) try { Marshal.ReleaseComObject(comObj); } catch {}
                    try { Marshal.ReleaseComObject(monikers[0]); } catch {}
                    monikers[0] = null;
                }
            }
        } catch (Exception ex) {
            result.Success = false;
            result.Error = ex.Message;
        } finally {
            if (enumMk != null) try { Marshal.ReleaseComObject(enumMk); } catch {}
            if (bindCtx != null) try { Marshal.ReleaseComObject(bindCtx); } catch {}
            if (rot != null) try { Marshal.ReleaseComObject(rot); } catch {}
        }
        return result;
    }
}
"@
    }

    $closedViaRot = $false
    try {
        $result = [RotManager]::CloseDatabaseIfOpen($resolved)
        if ($result.ClosedCount -gt 0) {
            Write-Status -Message ("Cerrada(s) {0} instancia(s) COM de la BD: {1}" -f $result.ClosedCount, $resolved) -Color Yellow
            $closedViaRot = $true
        }
        if ($result.Error) {
            Write-Status -Message ("ROT warning: {0}" -f $result.Error) -Color DarkYellow
        }
    } catch {
        # ROT no disponible — no es critico
    }

    # Fallback: si el ROT no cerro nada, buscar proceso MSACCESS con .laccdb bloqueado
    if (-not $closedViaRot) {
        $laccdb = [System.IO.Path]::ChangeExtension($resolved, ".laccdb")
        if (Test-Path -LiteralPath $laccdb) {
            Write-Status -Message ("Detectado lock activo: {0}" -f $laccdb) -Color Yellow

            # Buscar MSACCESS.EXE por CommandLine (contiene la ruta del .accdb que abrio)
            $dbFileName = [System.IO.Path]::GetFileName($resolved)
            $cimProcs = @(Get-CimInstance Win32_Process -Filter "Name = 'MSACCESS.EXE'" -ErrorAction SilentlyContinue)
            $killed = $false

            foreach ($cim in $cimProcs) {
                if ($cim.CommandLine -and $cim.CommandLine -match [regex]::Escape($dbFileName)) {
                    Write-Status -Message ("Cerrando MSACCESS PID {0} (CommandLine contiene: {1})" -f $cim.ProcessId, $dbFileName) -Color Yellow
                    try {
                        Stop-Process -Id $cim.ProcessId -Force -ErrorAction Stop
                        $killed = $true
                    } catch {
                        Write-Status -Message ("No se pudo cerrar MSACCESS PID {0}: {1}" -f $cim.ProcessId, $_.Exception.Message) -Color Red
                    }
                }
            }

            if (-not $killed -and $cimProcs.Count -gt 0) {
                Write-Status -Message ("Ningun MSACCESS contiene '{0}' en CommandLine. PIDs activos: {1}" -f $dbFileName, (($cimProcs | ForEach-Object { $_.ProcessId }) -join ', ')) -Color DarkYellow
            }

            if ($killed) {
                $timeout = 5; $elapsed = 0
                while ((Test-Path -LiteralPath $laccdb) -and ($elapsed -lt $timeout)) {
                    Start-Sleep -Milliseconds 500
                    $elapsed += 0.5
                }
                if (Test-Path -LiteralPath $laccdb) {
                    # Proceso muerto pero lock persiste — intentar borrar (seguro si el proceso ya no existe)
                    try {
                        Remove-Item -LiteralPath $laccdb -Force -ErrorAction Stop
                        Write-Status -Message "Lock huerfano eliminado." -Color Green
                    } catch {
                        Write-Status -Message ("No se pudo eliminar .laccdb huerfano: {0}" -f $_.Exception.Message) -Color Red
                    }
                } else {
                    Write-Status -Message "Lock liberado correctamente." -Color Green
                }
            }
        }
    }
}

function Get-AccessProcessId {
    # Detecta el PID del proceso MSACCESS asociado a una instancia COM.
    # Estrategia: hwnd -> GetWindowThreadProcessId. Fallback: diff de procesos pre/post.
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$AccessApp,
        [int[]]$PrePids = @()
    )

    # Intento 1: hWndAccessApp
    try {
        $hwnd = [IntPtr]$AccessApp.hWndAccessApp
        if ($hwnd -and $hwnd -ne [IntPtr]::Zero) {
            $pid = Get-ProcessIdFromHwnd -Hwnd $hwnd
            if ($pid -gt 0) { return $pid }
        }
    } catch {}

    # Intento 2: diff de procesos MSACCESS
    try {
        $post = @(Get-Process MSACCESS -ErrorAction SilentlyContinue | Select-Object -Property Id, StartTime)
        $new = @($post | Where-Object { $_.Id -notin $PrePids })
        if ($new.Count -ge 1) {
            $picked = $new | Sort-Object -Property StartTime -Descending | Select-Object -First 1
            return [int]$picked.Id
        }
    } catch {}

    return $null
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
        # Cerrar SOLO la instancia COM que tenga esta BD abierta, sin tocar otras instancias de Access
        Close-TargetAccessDbIfOpen -AccessPath $AccessPath

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

        $access.OpenCurrentDatabase($AccessPath, $false, $Password)
        try { $access.DoCmd.SetWarnings($false) } catch {}

        $accessPid = Get-AccessProcessId -AccessApp $access -PrePids $prePids

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

function Get-VbComponentNames {
    # Enumera nombres de componentes VBA exportables (BAS, CLS, Form).
    # Libera cada COM reference tras leer el nombre.
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject
    )
    $components = $VbProject.VBComponents
    $names = @()
    try {
        for ($i = 1; $i -le $components.Count; $i++) {
            $c = $components.Item($i)
            try {
                $ext = Get-ComponentExtension -Component $c -ModuleName $c.Name
                if ($ext) { $names += $c.Name }
            } finally {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($c) | Out-Null } catch {}
            }
        }
    } finally {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($components) | Out-Null } catch {}
    }
    return ($names | Sort-Object -Unique)
}

function Resolve-VbComponentName {
    # Resuelve el nombre de un componente en el VBProject.
    # Access internamente prefija los code-behind de formularios con "Form_",
    # pero los usuarios usan el nombre del formulario sin prefijo.
    # Intenta: $Name tal cual -> "Form_$Name"
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]$VbProject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    try {
        $c = $VbProject.VBComponents.Item($Name)
        if ($c) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($c) | Out-Null } catch {}
            return $Name
        }
    } catch {}

    $altName = "Form_" + $Name
    try {
        $c = $VbProject.VBComponents.Item($altName)
        if ($c) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($c) | Out-Null } catch {}
            return $altName
        }
    } catch {}

    return $Name  # devolver el original — el caller decidira si lanzar error
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
        $ModuleName = Resolve-VbComponentName -VbProject $VbProject -Name $ModuleName
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

            if (-not $AccessApplication) {
                # Sin sesion COM no es posible exportar la UI del formulario
                throw ("Se necesita -AccessApplication para exportar el formulario '{0}' con SaveAsText." -f $formName)
            }

            try {
                # acForm = 2
                $AccessApplication.SaveAsText(2, $formName, $tmp)
            } catch {
                throw ("SaveAsText lanzo excepcion para '{0}': {1}" -f $formName, $_.Exception.Message)
            }

            # Verificar integridad: SaveAsText puede completarse sin excepcion pero producir un archivo
            # incompleto si el formulario esta abierto en modo diseno o bloqueado internamente.
            # Un .form.txt valido siempre contiene la linea "Begin Form".
            $savedContent = $null
            if (Test-Path -Path $tmp) {
                try { $savedContent = Get-Content -Path $tmp -Raw -Encoding Default -ErrorAction Stop } catch {}
            }
            if (-not $savedContent -or $savedContent -notmatch 'Begin Form') {
                throw ("SaveAsText produjo un archivo incompleto para '{0}' (falta 'Begin Form'). " +
                       "Asegurate de que el formulario no este abierto en modo diseno en ninguna instancia de Access." -f $formName)
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

    # Nombres candidatos: el original y con prefijo Form_ (los usuarios no ponen Form_)
    $nameVariants = @($moduleNameText)
    if ($moduleNameText -notmatch '^Form_') {
        $nameVariants += ("Form_" + $moduleNameText)
    }

    $subFolders = @("forms", "classes", "modules", "")
    switch ($ImportMode) {
        "Form" { $extensions = @(".form.txt", ".frm") }
        "Code" { $extensions = @(".cls", ".bas") }
        default { $extensions = @(".form.txt", ".frm", ".cls", ".bas") }
    }

    # Extension tiene prioridad sobre carpeta, y nombre original sobre Form_ prefijado.
    foreach ($ext in $extensions) {
        foreach ($tryName in $nameVariants) {
            foreach ($folder in $subFolders) {
                $searchPath = if ($folder) { Join-Path -Path $modulesPathText -ChildPath $folder } else { $modulesPathText }
                if (-not (Test-Path -Path $searchPath)) { continue }

                $candidate = Join-Path -Path $searchPath -ChildPath ($tryName + $ext)
                if (Test-Path -Path $candidate) { return $candidate }
            }
        }
    }

    $any = Get-ChildItem -Path $modulesPathText -File -Recurse -Include "*.bas", "*.cls", "*.frm", "*.form.txt" -ErrorAction SilentlyContinue |
        Where-Object {
            $bn = $_.BaseName
            $fn = $_.Name -replace '\.form\.txt$', ''
            foreach ($tryName in $nameVariants) {
                if ($bn -ieq $tryName -or $fn -ieq $tryName) { return $true }
            }
            return $false
        } |
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
            # Cerrar el formulario si esta abierto — LoadFromText falla con "Cancelo la operacion anterior" si no
            try { $AccessApplication.DoCmd.Close(2, $formName, 1) } catch {}  # acForm=2, acSaveNo=1
            # acForm = 2
            $AccessApplication.LoadFromText(2, $formName, $tmpAnsi)
            return
        }

        # Determinar tipo esperado por extension del archivo fuente
        # vbext_ct_StdModule = 1 (.bas), vbext_ct_ClassModule = 2 (.cls)
        $extLower = $ext.ToLower()
        $expectedType = if ($extLower -eq '.cls') { 2 } else { 1 }

        # Strip metadatos VBE (VERSION, BEGIN/END, Attribute VB_*) del archivo temporal.
        # AddFromFile no los parsea y los inyectaria como codigo.
        Strip-VbaMetadataHeader -FilePath $tmpAnsi

        # Comprobar si el componente ya existe en el VBProject
        $component = $null
        $componentExists = $false
        try {
            $component = $VbProject.VBComponents.Item($ModuleName)
            $componentExists = $true
        } catch {
            $componentExists = $false
        }

        if ($componentExists) {
            $existingType = $component.Type  # 1=BAS, 2=CLS
            if ($existingType -ne $expectedType) {
                # Tipo cambia (ej: BAS->CLS o CLS->BAS) -- borrar y recrear
                # porque Access no permite cambiar el tipo de un componente existente
                Write-Status -Message ("Tipo cambia para '{0}': {1}->{2} -- Remove + Add" -f $ModuleName, $existingType, $expectedType) -Color Yellow
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component) | Out-Null } catch {}
                $component = $null
                $VbProject.VBComponents.Remove($VbProject.VBComponents.Item($ModuleName))
                $componentExists = $false
            }
        }

        if ($componentExists) {
            # Tipo correcto -- reemplazar codigo sin tocar el componente (evita dialogo VBE)
            $codeModule = $component.CodeModule
            $count = $codeModule.CountOfLines
            if ($count -gt 0) { $codeModule.DeleteLines(1, $count) }
            $codeModule.AddFromFile($tmpAnsi)
        } else {
            # Componente no existe (nuevo o recien borrado por cambio de tipo)
            # Crear con tipo explicito via Add(type) -- nunca Import que depende de headers VBE
            $component = $VbProject.VBComponents.Add($expectedType)
            $component.Name = $ModuleName
            $codeModule = $component.CodeModule
            $codeModule.AddFromFile($tmpAnsi)
        }

    } finally {
        if ($tmpAnsi -and (Test-Path -Path $tmpAnsi)) { Remove-Item -Path $tmpAnsi -Force -ErrorAction SilentlyContinue }
        foreach ($obj in @($codeModule, $component)) {
            if ($obj) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch {} }
        }
    }
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

        $tableDefs = $database.TableDefs
        $tables = @()
        for ($i = 0; $i -lt $tableDefs.Count; $i++) {
            $td = $tableDefs[$i]
            try {
                if ($td.Name -notmatch "^MSys" -and $td.Name -notmatch "^~") {
                    $tables += $td.Name
                }
            } catch {}
        }
        $tables = $tables | Sort-Object

        [void]$sb.AppendLine("## Tablas ($($tables.Count))")
        [void]$sb.AppendLine("")

        foreach ($tableName in $tables) {
            try {
                $td = $database.TableDefs[$tableName]
                [void]$sb.AppendLine("### $tableName")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("| Campo | Tipo | Tamano | Requerido | PK |")
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
                [void]$sb.AppendLine("_Error leyendo tabla: $tableName - $($_.Exception.Message)_")
                [void]$sb.AppendLine("")
            }
        }

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

        # FIX: renombrada $tdConnect para no sobreescribir $connect del scope exterior
        $linkedSources = @{}
        for ($i = 0; $i -lt $tableDefs.Count; $i++) {
            $td = $tableDefs[$i]
            try {
                $tdConnect = $td.Connect
                if (-not [string]::IsNullOrEmpty($tdConnect) -and $tdConnect -match ";DATABASE=(.+)$") {
                    $linkedDbPath = $Matches[1].Trim()
                    if (-not $linkedSources.ContainsKey($linkedDbPath)) {
                        $linkedSources[$linkedDbPath] = [System.Collections.Generic.List[string]]::new()
                    }
                    $linkedSources[$linkedDbPath].Add($td.Name)
                }
            } catch {}
        }

        $unreachableBackends = @($linkedSources.Keys | Where-Object { -not (Test-Path -Path $_) })
        if ($unreachableBackends.Count -gt 0) {
            [void]$sb.AppendLine("## Backends vinculados no alcanzados")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("Las siguientes bases de datos vinculadas no estaban disponibles al generar este ERD.")
            [void]$sb.AppendLine("Sus tablas aparecen en el listado de tablas pero su estructura no pudo verificarse.")
            [void]$sb.AppendLine("")
            foreach ($linkedPath in $unreachableBackends) {
                $linkedTables = $linkedSources[$linkedPath] -join ", "
                [void]$sb.AppendLine("- ``$linkedPath`` - tablas vinculadas: $linkedTables")
            }
            [void]$sb.AppendLine("")
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

    $names = @()

    if ($ModuleName -and $ModuleName.Count -gt 0) {
        $names = @($ModuleName | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } else {
        $names = @(Get-VbComponentNames -VbProject $VbProject)
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

$session = $null

try {
    $DestinationRoot = Resolve-DestinationRoot -DestinationRoot $DestinationRoot

    if ($Action -ne "Generate-ERD") {
        $AccessPath = Resolve-AccessPath -AccessPath $AccessPath
        $ModulesPath = Resolve-ModulesPath -DestinationRoot $DestinationRoot -AccessPath $AccessPath -Action $Action

        Write-Status -Message ("Accion: {0}" -f $Action) -Color Yellow
        Write-Status -Message ("Base de datos: {0}" -f $AccessPath) -Color Yellow
        Write-Status -Message ("Carpeta: {0}" -f $ModulesPath) -Color Yellow
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

        $targets = @()
        if ($normalizedModules.Count -gt 0) {
            $targets = $normalizedModules
        } else {
            $targets = @(Get-VbComponentNames -VbProject $vbProject)
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

    } elseif ($Action -eq "Generate-ERD") {
        if ([string]::IsNullOrWhiteSpace($BackendPath)) {
            $candidates = Get-ChildItem -Path (Get-Location) -File -Filter "*_Datos.accdb" -ErrorAction SilentlyContinue
            if (-not $candidates) {
                $candidates = Get-ChildItem -Path (Get-Location) -File -Filter "*_Datos.mdb" -ErrorAction SilentlyContinue
            }

            if ($candidates) {
                if ($candidates.Count -gt 1) {
                    Write-Status -Message "ADVERTENCIA: Multiples backends encontrados, usando el primero: $($candidates[0].Name)" -Color Yellow
                }
                $BackendPath = $candidates[0].FullName
            } else {
                throw "No se especifico -BackendPath y no se encontro ningun archivo *_Datos.accdb/.mdb en el directorio actual."
            }
        }

        $BackendPath = (Resolve-Path -Path $BackendPath).Path
        Write-Status -Message ("Backend: {0}" -f $BackendPath) -Color Yellow

        if ([string]::IsNullOrWhiteSpace($ErdPath)) {
            $parent = Split-Path -Parent $DestinationRoot
            $ErdPath = Join-Path -Path $parent -ChildPath "ERD"
        }

        if (-not (Test-Path -Path $ErdPath)) {
            New-Item -ItemType Directory -Force -Path $ErdPath | Out-Null
        }
        $ErdPath = (Resolve-Path -Path $ErdPath).Path
        Write-Status -Message ("ERD Folder: {0}" -f $ErdPath) -Color Yellow

        $backendName = [System.IO.Path]::GetFileNameWithoutExtension($BackendPath)
        $mdFile = Join-Path -Path $ErdPath -ChildPath ($backendName + ".md")

        Export-DataStructure -DatabasePath $BackendPath -OutputPath $mdFile -Password $Password

        Write-Status -Message ("OK ERD generado en: {0}" -f $mdFile) -Color Green

    } elseif ($Action -eq "Sandbox") {
        # =====================================================================
        # SANDBOX: Copiar backends vinculados al lado del frontend y
        #          revincular las tablas para que apunten a las copias locales.
        #          Resultado: un sandbox aislado de produccion.
        # =====================================================================
        $frontDir = Split-Path $AccessPath -Parent
        $bkpPassword = if ($BackendPassword) { $BackendPassword } else { $Password }

        # --- Fase 1: Descubrir backends vinculados via DAO ---
        Write-Status -Message "Descubriendo tablas vinculadas..." -Color Cyan
        $daoEngine = $null; $daoDb = $null
        $backendMap = @{}  # ruta_backend_original -> @(tabla1, tabla2, ...)
        try {
            $daoEngine = New-DaoDbEngine
            if (-not $daoEngine) { throw "No se pudo crear DAO.DBEngine" }
            $frontConn = if ($Password) { ";PWD=$Password" } else { "" }
            $daoDb = $daoEngine.OpenDatabase($AccessPath, $false, $false, $frontConn)
            $tdefs = $daoDb.TableDefs
            for ($i = 0; $i -lt $tdefs.Count; $i++) {
                $td = $tdefs[$i]
                $tName = $td.Name
                if ($tName -match "^~" -or $tName -match "^MSys") { continue }
                $srcTable = $td.SourceTableName
                $connect  = $td.Connect
                if ([string]::IsNullOrEmpty($srcTable)) { continue }
                # Extraer ruta del backend: ";DATABASE=C:\...\file.accdb;..."
                if ($connect -match "DATABASE=([^;]+)") {
                    $backendPath = $Matches[1].Trim()
                    if (-not $backendMap.ContainsKey($backendPath)) {
                        $backendMap[$backendPath] = @()
                    }
                    $backendMap[$backendPath] += $tName
                }
            }
        } finally {
            if ($daoDb) { try { $daoDb.Close() } catch {} }
            foreach ($o in @($daoDb, $daoEngine)) {
                if ($o) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null } catch {} }
            }
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        }

        if ($backendMap.Count -eq 0) {
            Write-Status -Message "El frontend no tiene tablas vinculadas. Nada que hacer." -Color Green
            return
        }

        $totalTables = ($backendMap.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        Write-Status -Message ("Encontrados {0} backend(s) con {1} tabla(s) vinculada(s):" -f $backendMap.Count, $totalTables) -Color Yellow
        foreach ($bk in $backendMap.Keys) {
            Write-Status -Message ("  {0} ({1} tablas)" -f $bk, $backendMap[$bk].Count) -Color Gray
        }

        # --- Fase 2: Copiar cada backend al lado del frontend ---
        Write-Status -Message "Copiando backends al directorio del frontend..." -Color Cyan
        $sidecarMap = @{}  # ruta_original -> ruta_sidecar
        foreach ($originalPath in $backendMap.Keys) {
            $backendFileName = Split-Path $originalPath -Leaf
            $sidecarPath = Join-Path $frontDir $backendFileName

            if ($sidecarPath -eq $originalPath) {
                Write-Status -Message ("  SKIP: {0} ya esta en el directorio del frontend" -f $backendFileName) -Color Yellow
                $sidecarMap[$originalPath] = $originalPath
                continue
            }

            if (Test-Path $sidecarPath) {
                Write-Status -Message ("  Reemplazando sidecar existente: {0}" -f $backendFileName) -Color Yellow
                Remove-Item $sidecarPath -Force
            }

            if (-not (Test-Path $originalPath)) {
                throw ("Backend no accesible: {0}" -f $originalPath)
            }

            Copy-Item -LiteralPath $originalPath -Destination $sidecarPath -Force
            Write-Status -Message ("  Copiado: {0}" -f $backendFileName) -Color Green
            $sidecarMap[$originalPath] = $sidecarPath
        }

        # --- Fase 3: Abrir frontend via COM y revincular ---
        # Esperar a que los locks se liberen
        Start-Sleep -Milliseconds 500

        Write-Status -Message "Abriendo frontend via COM para revincular tablas..." -Color Cyan
        $session = Open-AccessDatabase -AccessPath $AccessPath -Password $Password
        $comDb = $null

        try {
            $comDb = $session.AccessApplication.CurrentDb()

            # Guardar connect strings originales para rollback si falla la revinculacion
            $originalLinks = @{}
            for ($i = 0; $i -lt $comDb.TableDefs.Count; $i++) {
                $td = $comDb.TableDefs[$i]
                if (-not [string]::IsNullOrEmpty($td.SourceTableName)) {
                    $originalLinks[$td.Name] = @{
                        SourceTableName = $td.SourceTableName
                        Connect = $td.Connect
                    }
                }
            }

            # Eliminar todas las tablas vinculadas
            $toDelete = @($originalLinks.Keys)
            foreach ($tname in $toDelete) {
                try {
                    $comDb.TableDefs.Delete($tname)
                    Write-Status -Message ("  Desvinculada: {0}" -f $tname) -Color Gray
                } catch {
                    Write-Status -Message ("  WARN desvinculando {0}: {1}" -f $tname, $_.Exception.Message) -Color Yellow
                }
            }

            # Crear nuevos vinculos apuntando a los sidecars
            $okCount = 0; $errorCount = 0
            foreach ($originalPath in $backendMap.Keys) {
                $sidecarPath = $sidecarMap[$originalPath]
                $tables = $backendMap[$originalPath]
                $newConnect = ";DATABASE=$sidecarPath;PWD=$bkpPassword;"

                foreach ($tableName in $tables) {
                    try {
                        $newTd = $comDb.CreateTableDef($tableName, 0, $tableName, $newConnect)
                        $comDb.TableDefs.Append($newTd)
                        Write-Status -Message ("  OK: {0} -> {1}" -f $tableName, (Split-Path $sidecarPath -Leaf)) -Color Green
                        $okCount++
                    } catch {
                        Write-Status -Message ("  ERROR: {0} - {1}" -f $tableName, $_.Exception.Message) -Color Red
                        $errorCount++
                    }
                }
            }

            # Si hubo errores, intentar rollback restaurando vinculos originales
            if ($errorCount -gt 0) {
                Write-Status -Message "Errores detectados -- intentando rollback de vinculos originales..." -Color Yellow
                foreach ($tname in $originalLinks.Keys) {
                    # Solo restaurar las que no se revincularon exitosamente
                    $exists = $false
                    try { $null = $comDb.TableDefs($tname); $exists = $true } catch {}
                    if ($exists) { continue }
                    try {
                        $info = $originalLinks[$tname]
                        $restoreTd = $comDb.CreateTableDef($tname, 0, $info.SourceTableName, $info.Connect)
                        $comDb.TableDefs.Append($restoreTd)
                        Write-Status -Message ("  Restaurada: {0}" -f $tname) -Color DarkYellow
                    } catch {
                        Write-Status -Message ("  FALLO restaurar: {0} - {1}" -f $tname, $_.Exception.Message) -Color Red
                    }
                }
            }

            Write-Status -Message ("Sandbox completado: {0} OK, {1} errores" -f $okCount, $errorCount) -Color $(if ($errorCount -eq 0) { "Green" } else { "Yellow" })

        } finally {
            if ($comDb) {
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comDb) | Out-Null } catch {}
            }
        }

    } else {
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
