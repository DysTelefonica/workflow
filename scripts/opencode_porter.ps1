# --- CONFIGURACIÓN DE RUTAS ---
$paths = @{
    Config = "$env:USERPROFILE\.config\opencode"
    Data   = "$env:USERPROFILE\.local\share\opencode"
    Project = "$pwd" # Usa la carpeta actual donde ejecutes el script
}

function Show-Header {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "      OPENCODE - TRANSPORTADOR DE ALMAS        " -ForegroundColor White -BackgroundColor DarkCyan
    Write-Host "===============================================" -ForegroundColor Cyan
}

function Export-OpenCode {
    $zipName = "OpenCode_Transfer_$(Get-Date -Format 'yyyyMMdd').zip"
    $destZip = Join-Path ([Environment]::GetFolderPath("Desktop")) $zipName
    $temp = Join-Path $env:TEMP "OC_Export"
    
    if (Test-Path $temp) { Remove-Item -Recurse -Force $temp }
    New-Item -ItemType Directory -Path $temp | Out-Null

    Write-Host "`n[+] Preparando equipaje..." -ForegroundColor Yellow
    
    # Copia selectiva basada en rutas conocidas 
    if (Test-Path $paths.Config) { Copy-Item -Path $paths.Config -Destination (Join-Path $temp "config") -Recurse }
    if (Test-Path $paths.Data) { 
        New-Item -ItemType Directory -Path (Join-Path $temp "data") | Out-Null
        Copy-Item -Path (Join-Path $paths.Data "opencode.db") -Destination (Join-Path $temp "data")
        Copy-Item -Path (Join-Path $paths.Data "storage\session_diff") -Destination (Join-Path $temp "data") -Recurse
    }
    if (Test-Path (Join-Path $paths.Project "AGENTS.md")) { 
        Copy-Item (Join-Path $paths.Project "AGENTS.md") $temp 
    }

    Write-Host "[+] Creando archivo comprimido en el Escritorio..." -ForegroundColor Yellow
    Compress-Archive -Path "$temp\*" -DestinationPath $destZip -Update
    Remove-Item -Recurse -Force $temp
    
    Write-Host "`n[!] ÉXITO: Llévate '$zipName' a la oficina." -ForegroundColor Green
    Pause
}

function Import-OpenCode {
    $zipPath = Read-Host "`nArrastra el archivo ZIP aquí y pulsa Enter"
    $zipPath = $zipPath.Trim('"')

    if (-not (Test-Path $zipPath)) {
        Write-Host "[-] Error: No encuentro el archivo." -ForegroundColor Red
        Pause
        return
    }

    Write-Host "`n[!] ADVERTENCIA: Se sobrescribirá la config actual. ¿Continuar? (S/N)" -ForegroundColor Red
    $confirm = Read-Host
    if ($confirm -ne "S") { return }

    $temp = Join-Path $env:TEMP "OC_Import"
    if (Test-Path $temp) { Remove-Item -Recurse -Force $temp }
    Expand-Archive -Path $zipPath -DestinationPath $temp

    Write-Host "[+] Restaurando sistema..." -ForegroundColor Yellow
    
    if (Test-Path (Join-Path $temp "config")) { Copy-Item (Join-Path $temp "config\*") $paths.Config -Recurse -Force }
    if (Test-Path (Join-Path $temp "data")) { Copy-Item (Join-Path $temp "data\*") $paths.Data -Recurse -Force }
    if (Test-Path (Join-Path $temp "AGENTS.md")) { Copy-Item (Join-Path $temp "AGENTS.md") $paths.Project -Force }

    Remove-Item -Recurse -Force $temp
    Write-Host "`n[!] IMPORTACIÓN COMPLETADA. Reinicia OpenCode." -ForegroundColor Green
    Pause
}

# --- BUCLE PRINCIPAL ---
do {
    Show-Header
    Write-Host "Ruta actual del proyecto: $($paths.Project)" -ForegroundColor Gray
    Write-Host "`n1. EXPORTAR para la oficina (Crear ZIP)"
    Write-Host "2. IMPORTAR en este PC (Leer ZIP)"
    Write-Host "3. Salir"
    
    $choice = Read-Host "`nSelecciona una opción"

    switch ($choice) {
        "1" { Export-OpenCode }
        "2" { Import-OpenCode }
        "3" { exit }
    }
} while ($true)