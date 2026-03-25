# ============================================================
# sync-opencode.ps1 — Sincroniza configuración de OpenCode
# entre máquinas usando engram sync y Git
# ============================================================
# Estructura del repo:
#   profiles/
#   ├── casa/         ← configs de casa (adm1)
#   │   ├── opencode/
#   │   └── engram/  ← chunks JSON de casa
#   └── oficina/      ← configs de oficina (adm.defensa)
#       ├── opencode/
#       └── engram/  ← chunks JSON de oficina
#   bin/              ← wrapper obsidian (común)
# ============================================================
# Uso:
#   .\sync-opencode.ps1 -Profile casa     # Sincronizar casa
#   .\sync-opencode.ps1 -Profile oficina  # Sincronizar oficina
#   .\sync-opencode.ps1 -Action push     # Push (perfil auto)
#   .\sync-opencode.ps1 -Action pull     # Pull (perfil auto)
#   .\sync-opencode.ps1 -Status         # Ver estado
# ============================================================

param(
    [ValidateSet("push", "pull", "status")]
    [string]$Action = "status",
    [ValidateSet("casa", "oficina", "auto")]
    [string]$Profile = "auto"
)

# --- CARGAR CONFIGURACIÓN ---
$CONFIG_FILE = "$PSScriptRoot/sync-config.json"
if (-not (Test-Path $CONFIG_FILE)) {
    Write-Host "[ERR] No se encontro sync-config.json" -ForegroundColor Red
    exit 1
}

$CONFIG = Get-Content $CONFIG_FILE | ConvertFrom-Json

# --- DETECTAR PERFIL AUTO ---
if ($Profile -eq "auto") {
    $currentUser = $env:USERNAME
    if ($currentUser -eq "adm1") { $Profile = "casa" }
    elseif ($currentUser -eq "adm.defensa") { $Profile = "oficina" }
    else { $Profile = "casa" }
    Write-Host "   Perfil: $Profile ($currentUser)" -ForegroundColor Gray
}

$PERFIL = $CONFIG.profiles.$Profile
if (-not $PERFIL) {
    Write-Host "[ERR] Perfil '$Profile' no encontrado" -ForegroundColor Red
    exit 1
}

# --- VARIABLES ---
$OPENCODE_DIR = $PERFIL.paths.opencode_dir
$ENGRAM_DATA_DIR = $PERFIL.paths.engram_data_dir
$OBSIDIAN_WRAPPER = $CONFIG.common.obsidian_wrapper
$REPO_NAME = $CONFIG.repo.name
$REPO_ORG = $CONFIG.repo.org

$REPO_DIR = Split-Path -Parent $PSScriptRoot
$PROFILE_DIR = "$REPO_DIR/profiles/$Profile"
$OPENCODE_DEST = "$PROFILE_DIR/opencode"
$ENGRAM_DEST = "$PROFILE_DIR/engram"
$BIN_DEST = "$REPO_DIR/bin"

# --- FUNCIONES ---
function Write-Step { param([string]$M) { Write-Host "`n>>> $M" -ForegroundColor Cyan } }
function Write-Success { param([string]$M) { Write-Host "   [OK] $M" -ForegroundColor Green } }
function Write-Err { param([string]$M) { Write-Host "   [ERR] $M" -ForegroundColor Red } }
function Ensure-Dir { param([string]$P) { if (-not (Test-Path $P)) { New-Item -ItemType Directory -Path $P -Force | Out-Null } } }
function Get-Chunks { param([string]$D) if (Test-Path $D) { Get-ChildItem $D -Filter "*.json" -ErrorAction SilentlyContinue } else { @() } }

# --- STATUS ---
function Show-Status {
    Write-Host "`n=== Estado ===" -ForegroundColor Yellow
    Write-Host "   Perfil: $Profile | Repo: $REPO_ORG/$REPO_NAME" -ForegroundColor Gray
    
    $localChunks = Get-Chunks $ENGRAM_DATA_DIR
    $repoChunks = Get-Chunks $ENGRAM_DEST
    
    Write-Host "`n--- Engram Local ($Profile) ---" -ForegroundColor Magenta
    if ($localChunks) { Write-Host "  $($localChunks.Count) chunks locales" -ForegroundColor White }
    else { Write-Host "  Sin chunks locales" -ForegroundColor Gray }
    
    Write-Host "`n--- Engram Repo ($Profile) ---" -ForegroundColor Magenta
    if ($repoChunks) { Write-Host "  $($repoChunks.Count) chunks en repo" -ForegroundColor White }
    else { Write-Host "  Sin chunks en repo" -ForegroundColor Gray }
    
    Write-Host "`n--- OpenCode Local ---" -ForegroundColor Magenta
    if (Test-Path $OPENCODE_DIR) { Write-Host "  OK: $OPENCODE_DIR" -ForegroundColor Green }
    else { Write-Host "  NO ENCONTRADO" -ForegroundColor Red }
    
    Write-Host "`n--- Obsidian Wrapper ---" -ForegroundColor Magenta
    if (Test-Path $OBSIDIAN_WRAPPER) { Write-Host "  OK" -ForegroundColor Green }
    else { Write-Host "  NO ENCONTRADO" -ForegroundColor Red }
    
    Write-Host "`n=========================================" -ForegroundColor Yellow
    Write-Host "`n.\sync-opencode.ps1 -Profile $Profile -Action push" -ForegroundColor Cyan
    Write-Host ".\sync-opencode.ps1 -Profile $Profile -Action pull" -ForegroundColor Cyan
}

# --- PUSH ---
function Push-Changes {
    Set-Location $REPO_DIR
    Write-Step "Push $Profile → repo..."
    
    # Engram sync
    Write-Host "`n--- Engram ---" -ForegroundColor Magenta
    Ensure-Dir $ENGRAM_DEST
    $tempDir = "$ENGRAM_DATA_DIR\sync-temp"
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    Ensure-Dir $tempDir
    $env:ENGRAM_CHUNK_DIR = $tempDir
    engram sync 2>&1 | Out-Null
    $chunks = Get-Chunks $tempDir
    if ($chunks) {
        foreach ($c in $chunks) {
            Copy-Item $c.FullName $ENGRAM_DEST -Force
            Write-Host "    + $($c.Name)" -ForegroundColor Green
        }
        Write-Success "Chunks exportados"
    } else { Write-Host "  Sin nuevos chunks" -ForegroundColor Gray }
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item Env:ENGRAM_CHUNK_DIR -ErrorAction SilentlyContinue
    
    # OpenCode
    Write-Host "`n--- OpenCode ---" -ForegroundColor Magenta
    if (Test-Path $OPENCODE_DIR) {
        if (Test-Path $OPENCODE_DEST) { Remove-Item $OPENCODE_DEST -Recurse -Force }
        Ensure-Dir $OPENCODE_DEST
        Copy-Item "$OPENCODE_DIR/*" $OPENCODE_DEST -Recurse -Force -Exclude @("node_modules",".git","cache","*.log")
        Write-Success "opencode/ copiado"
    } else { Write-Err "opencode no encontrado: $OPENCODE_DIR" }
    
    # Obsidian wrapper (común)
    Write-Host "`n--- Obsidian Wrapper ---" -ForegroundColor Magenta
    Ensure-Dir $BIN_DEST
    if (Test-Path $OBSIDIAN_WRAPPER) {
        Copy-Item $OBSIDIAN_WRAPPER "$BIN_DEST/" -Force
        Write-Success "bin/ copiado"
    } else { Write-Err "Wrapper no encontrado" }
    
    # Archivos de sync
    Copy-Item $CONFIG_FILE "$REPO_DIR/" -Force
    Copy-Item "$PSCommandPath" "$REPO_DIR/scripts/" -Force
    
    # Git
    Write-Step "Git..."
    git add -A
    $status = git status --short
    if ($status) {
        git commit -m "sync [$Profile]: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        Write-Success "Commit creado"
        git push origin main
        Write-Success "Push completado!"
    } else { Write-Host "  Sin cambios" -ForegroundColor Gray }
}

# --- PULL ---
function Pull-Changes {
    Set-Location $REPO_DIR
    Write-Step "Pull repo → $Profile..."
    
    git pull origin main
    Write-Success "Pull completado"
    
    # Engram import
    Write-Host "`n--- Engram ---" -ForegroundColor Magenta
    $chunks = Get-Chunks $ENGRAM_DEST
    if ($chunks) {
        foreach ($c in $chunks) {
            engram import $c.FullName 2>&1 | Out-Null
            Write-Host "    + $($c.Name)" -ForegroundColor Green
        }
    } else { Write-Host "  Sin chunks para importar" -ForegroundColor Gray }
    
    # OpenCode
    Write-Step "Restaurando OpenCode..."
    if (Test-Path $OPENCODE_DEST) {
        if (-not (Test-Path $OPENCODE_DIR)) { Ensure-Dir $OPENCODE_DIR }
        Copy-Item "$OPENCODE_DEST/*" $OPENCODE_DIR -Recurse -Force
        Write-Success "opencode/ restaurado"
    }
    
    # Obsidian wrapper
    Write-Step "Restaurando Obsidian Wrapper..."
    if (Test-Path "$BIN_DEST/obsidian") {
        Ensure-Dir (Split-Path $OBSIDIAN_WRAPPER)
        Copy-Item "$BIN_DEST/obsidian" $OBSIDIAN_WRAPPER -Force
        Write-Success "bin/ restaurado"
    }
    
    Write-Success "`nSincronización completa!"
    Write-Host "`nNOTA: Reiniciá OpenCode." -ForegroundColor Yellow
}

# --- MAIN ---
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  sync-opencode.ps1" -ForegroundColor White
Write-Host "  Repo: $REPO_ORG/$REPO_NAME | Perfil: $Profile" -ForegroundColor Gray
Write-Host "========================================" -ForegroundColor Cyan

switch ($Action) {
    "status" { Show-Status }
    "push"   { Push-Changes }
    "pull"   { Pull-Changes }
}
