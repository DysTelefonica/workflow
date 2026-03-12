# ==============================================================================
# VBA-SDD Framework — Script de Despliegue v4.0
# Fuente: carpeta local del repo de skills
#
# ESTRUCTURA ESPERADA EN EL REPO DE SKILLS:
#   rules/
#     engram-memory-quality.md
#     user_rules.md
#   sdd-protocol/SKILL.md
#   spec-writer/SKILL.md
#   spec-writer/references/spec_template.md
#   prd-writer/SKILL.md
#   prd-writer/references/prd_template.md
#   prd-writer/references/project_context_template.md
#   hotfix/SKILL.md
#   diario-sesion/SKILL.md
#   diario-sesion/references/diario_template.md
#   rfc-writer/SKILL.md  (opcional)
#   access-vba-sync/SKILL.md + cli.js + handler.js + VBAManager.ps1 + package.json
#   templates/
#     AGENTS_template.md
#
# USO:
#   .\deploy.ps1              → Proyecto nuevo desde cero
#   .\deploy.ps1 -UpdateOnly  → Sobreescribir framework en proyecto existente
# ==============================================================================

param(
    [switch]$UpdateOnly
)

# ==============================================================================
# CONFIGURACIÓN DE LA CARPETA FUENTE LOCAL
# Se guarda en %USERPROFILE%\.vba-sdd-source — se pregunta solo la primera vez
# ==============================================================================
$SourceConfigFile = "$env:USERPROFILE\.vba-sdd-source"

if (Test-Path $SourceConfigFile) {
    $SourcePath = Get-Content $SourceConfigFile -Raw | ConvertFrom-Json | Select-Object -ExpandProperty SourcePath
    if (!(Test-Path $SourcePath)) {
        Write-Host "`n  [!] La carpeta fuente ya no existe: $SourcePath" -ForegroundColor Yellow
        Write-Host "       Se pedirá una nueva ruta.`n" -ForegroundColor Yellow
        Remove-Item $SourceConfigFile -Force
        $SourcePath = $null
    } else {
        Write-Host "  Fuente: $SourcePath" -ForegroundColor Gray
    }
}

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    Write-Host "`n--- 📦 CONFIGURACIÓN DE FUENTE LOCAL ---" -ForegroundColor Yellow
    Write-Host "  Ruta de tu repo local de skills" -ForegroundColor Gray
    Write-Host "  (Solo se pregunta una vez)`n" -ForegroundColor Gray

    do {
        $SourcePath = Read-Host "Ruta completa (ej: C:\Dev\vba-sdd-skills)"
        $SourcePath = $SourcePath.Trim('"').Trim("'").TrimEnd('\')
        if (!(Test-Path $SourcePath)) {
            Write-Host "  [✗] No existe: $SourcePath" -ForegroundColor Red
            $SourcePath = $null
        }
    } while ([string]::IsNullOrWhiteSpace($SourcePath))

    @{ SourcePath = $SourcePath } | ConvertTo-Json | Set-Content $SourceConfigFile -Encoding UTF8
    Write-Host "  [✓] Guardado en $SourceConfigFile" -ForegroundColor Green
}

# ==============================================================================
# FUNCIÓN DE COPIA
# $RepoRel  : ruta relativa dentro del repo de skills
# $LocalDest: ruta de destino en el proyecto
# ==============================================================================
function Copy-SF($RepoRel, $LocalDest) {
    $Origin = Join-Path $SourcePath $RepoRel
    if (!(Test-Path $Origin)) {
        Write-Host "  [✗] No encontrado: $RepoRel" -ForegroundColor Red
        return
    }
    try {
        $Dir = Split-Path -Parent $LocalDest
        if ($Dir -and !(Test-Path $Dir)) {
            New-Item -ItemType Directory -Path $Dir -Force | Out-Null
        }
        Copy-Item -Path $Origin -Destination $LocalDest -Force
        Write-Host "  [✓] $LocalDest" -ForegroundColor Gray
    } catch {
        Write-Host "  [✗] Error: $RepoRel → $LocalDest" -ForegroundColor Red
    }
}

# ==============================================================================
# FUNCIÓN PRINCIPAL: copia todo el framework al proyecto destino
# $SD = SkillsDir destino en el proyecto  (ej: .trae/skills)
# $RD = RulesDir destino en el proyecto   (ej: .trae/rules)
# ==============================================================================
function Deploy-Framework($SD, $RD) {

    Write-Host "`n--- 📜 RULES ---" -ForegroundColor Yellow
    Copy-SF "rules/engram-memory-quality.md"  "$RD/engram-memory-quality.md"
    Copy-SF "rules/user_rules.md"             "$RD/user_rules.md"

    Write-Host "`n--- 🧠 SKILLS ---" -ForegroundColor Yellow

    # sdd-protocol
    Copy-SF "skills/sdd-protocol/SKILL.md"   "$SD/sdd-protocol/SKILL.md"

    # spec-writer
    Copy-SF "skills/spec-writer/SKILL.md"                        "$SD/spec-writer/SKILL.md"
    Copy-SF "skills/spec-writer/references/spec_template.md"     "$SD/spec-writer/references/spec_template.md"

    # prd-writer
    Copy-SF "skills/prd-writer/SKILL.md"                                   "$SD/prd-writer/SKILL.md"
    Copy-SF "skills/prd-writer/references/prd_template.md"                 "$SD/prd-writer/references/prd_template.md"
    Copy-SF "skills/prd-writer/references/project_context_template.md"     "$SD/prd-writer/references/project_context_template.md"

    # hotfix
    Copy-SF "skills/hotfix/SKILL.md"   "$SD/hotfix/SKILL.md"

    # diario-sesion
    Copy-SF "skills/diario-sesion/SKILL.md"                        "$SD/diario-sesion/SKILL.md"
    Copy-SF "skills/diario-sesion/references/diario_template.md"   "$SD/diario-sesion/references/diario_template.md"

    # rfc-writer (opcional)
    if (Test-Path (Join-Path $SourcePath "skills/rfc-writer/SKILL.md")) {
        Copy-SF "skills/rfc-writer/SKILL.md"   "$SD/rfc-writer/SKILL.md"
    } else {
        Write-Host "  [·] rfc-writer no encontrado — omitido" -ForegroundColor DarkGray
    }

    # access-vba-sync
    Write-Host "`n--- 🔄 ACCESS-VBA-SYNC ---" -ForegroundColor Yellow
    @(
        "SKILL.md",
        "cli.js",
        "handler.js",
        "VBAManager.ps1",
        "package.json"
    ) | ForEach-Object {
        Copy-SF "skills/access-vba-sync/$_"   "$SD/access-vba-sync/$_"
    }

    Write-Host "`n--- 📄 TEMPLATES ---" -ForegroundColor Yellow
    Copy-SF "templates/AGENTS_template.md"   "docs/templates/AGENTS_template.md"
}

# ==============================================================================
# MODO: UPDATE ONLY
# Sobreescribe rules + skills + templates en un proyecto existente
# No toca AGENTS.md ni project_context.md (son específicos del proyecto)
# ==============================================================================
if ($UpdateOnly) {

    Write-Host "`n🔄 VBA-SDD — ACTUALIZACIÓN DE FRAMEWORK" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan

    if (Test-Path ".trae/skills") {
        $SD = ".trae/skills"; $RD = ".trae/rules"
        Write-Host "  IDE: Trae" -ForegroundColor Gray
    } elseif (Test-Path "skills") {
        $SD = "skills"; $RD = ".rules"
        Write-Host "  IDE: Estándar" -ForegroundColor Gray
    } else {
        Write-Host "`n  [!] No se detectó estructura de skills." -ForegroundColor Yellow
        Write-Host "       Ejecuta .\deploy.ps1 sin -UpdateOnly para proyecto nuevo." -ForegroundColor Yellow
        exit
    }

    if (!(Test-Path "AGENTS.md")) {
        Write-Host "`n  [!] No se encontró AGENTS.md." -ForegroundColor Yellow
        Write-Host "       Ejecuta .\deploy.ps1 sin -UpdateOnly para proyecto nuevo." -ForegroundColor Yellow
        exit
    }

    Deploy-Framework $SD $RD

    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "✅ FRAMEWORK ACTUALIZADO" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "  Rules     : $RD (sobreescritas)" -ForegroundColor White
    Write-Host "  Skills    : $SD (sobreescritas)" -ForegroundColor White
    Write-Host "  Templates : docs/templates/ (sobreescritas)" -ForegroundColor White
    Write-Host "  AGENTS.md            : sin cambios" -ForegroundColor White
    Write-Host "  project_context.md   : sin cambios" -ForegroundColor White
    Write-Host "`n  Reinicia Trae para cargar las skills actualizadas.`n" -ForegroundColor Yellow
    exit
}

# ==============================================================================
# MODO: DEPLOY COMPLETO — proyecto nuevo
# ==============================================================================

Write-Host "`n🛠️  VBA-SDD — PROYECTO NUEVO" -ForegroundColor Cyan
Write-Host "==============================`n"

# 1. Detección del .accdb
$Files    = Get-ChildItem -Path (Get-Location) -File
$Frontend = $Files | Where-Object { $_.Extension -match "acc" -and $_.Name -notmatch "_[Dd]atos" } | Select-Object -First 1

if ($null -eq $Frontend) {
    Write-Host "⚠️  No se encontró .accdb en el directorio actual." -ForegroundColor Yellow
    $ok = Read-Host "¿Continuar igualmente? (S/N)"
    if ($ok -notmatch "S|s") { exit }
} else {
    Write-Host "✅ Detectado: $($Frontend.Name)" -ForegroundColor Green
}

# 2. Datos del proyecto
Write-Host "`n--- 📋 DATOS DEL PROYECTO ---" -ForegroundColor Yellow
$ProjectName   = Read-Host "Nombre del proyecto  (ej: BRASS)"
$ProjectStack  = Read-Host "Stack tecnológico    (ej: Access + VBA)"
$ProjectDomain = Read-Host "Dominio              (ej: Gestión de mantenimiento técnico)"
$ProjectPhase  = Read-Host "Fase actual          (ej: Autodescubrimiento)"

# 3. IDE
Write-Host "`n--- 🤖 IDE ---" -ForegroundColor Yellow
Write-Host "  1 → Trae   (.trae/rules + .trae/skills)" -ForegroundColor Gray
Write-Host "  2 → Cursor / VS Code / Estándar (.rules + skills)" -ForegroundColor Gray
$Choice = Read-Host "IDE (1 / 2)"

if ($Choice -eq "1") {
    $SD = ".trae/skills"; $RD = ".trae/rules"
} else {
    $SD = "skills"; $RD = ".rules"
}
Write-Host "  → Rules: $RD  |  Skills: $SD" -ForegroundColor Gray

# 4. Estructura de carpetas del proyecto
Write-Host "`n--- 📂 CARPETAS ---" -ForegroundColor Yellow
@(
    $RD,
    "$SD/sdd-protocol",
    "$SD/spec-writer/references",
    "$SD/prd-writer/references",
    "$SD/hotfix",
    "$SD/diario-sesion/references",
    "$SD/rfc-writer",
    "$SD/access-vba-sync",
    "docs/PRD",
    "docs/specs/active",
    "docs/specs/completed",
    "docs/templates",
    "references",
    "src/clases",
    "src/modulos",
    "src/formularios",
    ".engram"
) | ForEach-Object {
    if (!(Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
        Write-Host "  [+] $_" -ForegroundColor Gray
    }
}

# 5. Copiar framework
Deploy-Framework $SD $RD

# 6. npm install en access-vba-sync
Write-Host "`n--- 📦 ACCESS-VBA-SYNC (npm install) ---" -ForegroundColor Yellow
$SyncDir = "$SD/access-vba-sync"
if (Test-Path "$SyncDir/package.json") {
    Push-Location $SyncDir
    npm install --silent 2>&1 | Out-Null
    Pop-Location
    Write-Host "  [✓] Dependencias instaladas en $SyncDir" -ForegroundColor Green
} else {
    Write-Host "  [!] package.json no encontrado — omitido" -ForegroundColor Yellow
}

# 7. Export inicial de módulos VBA
Write-Host "`n--- 📤 EXPORT INICIAL (access-vba-sync start) ---" -ForegroundColor Yellow
if ($null -ne $Frontend -and (Test-Path "$SyncDir/cli.js")) {
    node "$SyncDir/cli.js" start --access "$($Frontend.FullName)"
    Write-Host "  [✓] Módulos exportados a src/" -ForegroundColor Green
} else {
    Write-Host "  [·] Omitido (no hay .accdb o access-vba-sync no instalado)" -ForegroundColor DarkGray
    Write-Host "       Ejecuta manualmente: node $SyncDir/cli.js start" -ForegroundColor DarkGray
}

# 8. Generar ERD
Write-Host "`n--- 🗂️  ERD (access-vba-sync generate-erd) ---" -ForegroundColor Yellow
$Backend = $Files | Where-Object { $_.Extension -match "acc" -and $_.Name -match "_[Dd]atos" } | Select-Object -First 1
if ($null -ne $Backend -and (Test-Path "$SyncDir/cli.js")) {
    node "$SyncDir/cli.js" generate-erd --backend "$($Backend.FullName)"
    Write-Host "  [✓] ERD generado en ERD/" -ForegroundColor Green
} elseif ($null -eq $Backend) {
    Write-Host "  [·] No se encontró *_Datos.accdb — omitido" -ForegroundColor DarkGray
    Write-Host "       Ejecuta manualmente: node $SyncDir/cli.js generate-erd --backend <ruta>" -ForegroundColor DarkGray
} else {
    Write-Host "  [·] access-vba-sync no disponible — omitido" -ForegroundColor DarkGray
}

# 9. Generar AGENTS.md
Write-Host "`n--- 🤖 AGENTS.MD ---" -ForegroundColor Yellow
$AgentsTemplate = "docs/templates/AGENTS_template.md"

if (Test-Path $AgentsTemplate) {
    $c = Get-Content $AgentsTemplate -Raw -Encoding UTF8
    $c = $c -replace "\{\{PROJECT_NAME\}\}", $ProjectName
    $c = $c -replace "\{\{STACK\}\}",        $ProjectStack
    $c = $c -replace "\{\{DOMAIN\}\}",       $ProjectDomain
    $c = $c -replace "\{\{PHASE\}\}",        $ProjectPhase
    $c = $c -replace "\{\{SKILLS_DIR\}\}",   $SD
    $c = $c -replace "\{\{RULES_DIR\}\}",    $RD
    Set-Content -Path "AGENTS.md" -Value $c -Encoding UTF8
    Write-Host "  [✓] AGENTS.md generado para '$ProjectName'" -ForegroundColor Green
} else {
    Write-Host "  [✗] AGENTS_template.md no encontrado en docs/templates/" -ForegroundColor Red
    Write-Host "       Comprueba que existe en: $SourcePath\templates\AGENTS_template.md" -ForegroundColor Yellow
}

# 10. Crear project_context.md desde plantilla (solo si no existe)
Write-Host "`n--- 📋 PROJECT_CONTEXT.MD ---" -ForegroundColor Yellow
$CtxTemplate = "$SD/prd-writer/references/project_context_template.md"
$CtxDest     = "references/project_context.md"

if (!(Test-Path $CtxDest)) {
    if (Test-Path $CtxTemplate) {
        $c = Get-Content $CtxTemplate -Raw -Encoding UTF8
        $c = $c -replace "\{NOMBRE_PROYECTO\}", $ProjectName
        Set-Content -Path $CtxDest -Value $c -Encoding UTF8
        Write-Host "  [✓] Creado desde plantilla — el agente lo completará en autodescubrimiento" -ForegroundColor Green
    } else {
        Write-Host "  [!] project_context_template.md no encontrado — omitido" -ForegroundColor Yellow
    }
} else {
    Write-Host "  [→] Ya existe — sin cambios" -ForegroundColor Gray
}

# 11. Prompt de arranque
Write-Host "`n--- 🔍 ARRANQUE ---" -ForegroundColor Yellow
$gen = Read-Host "¿Mostrar prompt de arranque del autodescubrimiento? (S/N)"

if ($gen -match "S|s") {
    $EntryForm = Read-Host "Formulario de entrada (ej: Form_FormInicio)"
    if ([string]::IsNullOrWhiteSpace($EntryForm)) { $EntryForm = "Form_FormInicio" }

    Write-Host "`n📋 COPIA ESTE PROMPT AL AGENTE:" -ForegroundColor Green
    Write-Host "-----------------------------------------------------------" -ForegroundColor Green
    Write-Host @"
Empieza el autodescubrimiento del proyecto $ProjectName.
El formulario de entrada es $EntryForm.

Sigue la FASE 0 de AGENTS.md:
1. mem_session_start + mem_context
2. Leer $SD/prd-writer/SKILL.md y referencias base
3. Leer src/formularios/$EntryForm.frm.txt y Form_$EntryForm.cls
4. Presentarme el árbol de dependencias antes de generar ningún PRD
"@
    Write-Host "-----------------------------------------------------------`n" -ForegroundColor Green
}

# Resumen final
Write-Host "`n==============================" -ForegroundColor Cyan
Write-Host "🏁 LISTO — $ProjectName" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host "  IDE    : $(if ($Choice -eq '1') {'Trae'} else {'Estándar'})" -ForegroundColor White
Write-Host "  Rules  : $RD" -ForegroundColor White
Write-Host "  Skills : $SD" -ForegroundColor White
Write-Host "  Fuente : $SourcePath" -ForegroundColor White
Write-Host "`n  Próximos pasos:" -ForegroundColor Yellow
Write-Host "  1. Revisa AGENTS.md" -ForegroundColor Gray
Write-Host "  2. Reinicia Trae (carga rules y skills)" -ForegroundColor Gray
Write-Host "  3. Verifica src/ y ERD/ generados" -ForegroundColor Gray
Write-Host "  4. Copia el prompt de arranque al agente`n" -ForegroundColor Gray
