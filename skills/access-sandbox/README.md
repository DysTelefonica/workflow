# access-sandbox — Provisioning de Workspaces Access Locales

Skill agnóstico para preparar un workspace Access local y autocontenido a partir de un backend fuente.
Sin rutas hardcodeadas, sin código específico de proyecto.

> **Rol:** Esta skill *prepara* el sandbox. Para leer, consultar o modificar datos dentro del sandbox,
> usar **`access-query`** sobre el sandbox ya localizeado.

---

## Archivos

| Archivo | Descripción |
|---|---|
| `SandboxManager.ps1` | Motor de provisioning (PowerShell) |
| `SKILL.md` | Especificación técnica completa (referencia canonical) |
| `README.md` | Este archivo — guía de uso rápida |

---

## Quick start

```powershell
# Invocación directa (ruta absoluta al script)
& "C:\Users\adm\.config\opencode\.agent\skills\access-sandbox\SandboxManager.ps1" `
    -Action New-Sandbox `
    -SourceBackend "C:\proyecto\datos\MiBackend.accdb" `
    -SandboxPath "C:\sandbox\MiBackend.accdb"
```

---

## Acciones disponibles

### 1. New-Sandbox

Copia un backend fuente a una ruta de sandbox.

```powershell
-Action New-Sandbox
-SourceBackend "C:\proyecto\datos\Backend.accdb"
-SandboxPath "C:\sandbox\Backend.accdb"
```

**No usa COM** — `Copy-Item` directo (más rápido y seguro para crear la copia inicial).

---

### 2. Discover-LinkedTables

Abre el sandbox headless y retorna las tablas vinculadas.

```powershell
-Action Discover-LinkedTables
-SandboxPath "C:\sandbox\Backend.accdb"
```

**Retorna array de objetos** (pipeline output):

```powershell
TableName    : tblContratos
Connect      : ;DATABASE=C:\proyecto\datos\Backend.accdb
SourceTable  : tblContratos
IsLinked     : True
SourcePath   : C:\proyecto\datos\Backend.accdb
```

---

### 3. Make-Sidecar

Crea una copia sin password de un backend protegido. Útil cuando `Localize-Sandbox` necesita acceder a datos protegidos.

```powershell
-Action Make-Sidecar
-SourceBackend "C:\proyecto\datos\Backend_protegido.accdb"
-Password "secret"
-SidecarSuffix "_nopass"   # genera Backend_protegido_nopass.accdb
```

**Genera:** `C:\proyecto\datos\Backend_protegido_nopass.accdb`

---

### 4. Localize-Sandbox

Reemplaza todas las tablas vinculadas en el sandbox con tablas locales (datos copiados desde el source).

```powershell
# Single-backend: usando sidecar (sin password)
-Action Localize-Sandbox
-SandboxPath "C:\sandbox\Backend.accdb"
-SourceSidecar "C:\proyecto\datos\Backend_nopass.accdb"

# Multi-backend: mapa de sidecars por backend
$map = @{
    "C:\proyecto\datos\Backend1.accdb" = "C:\proyecto\datos\Backend1_nopass.accdb"
    "C:\proyecto\datos\Backend2.accdb" = "C:\proyecto\datos\Backend2_nopass.accdb"
} | ConvertTo-Json -Compress

-Action Localize-Sandbox
-SandboxPath "C:\sandbox\Backend.accdb"
-BackendSidecarMapJson $map

# O usando backend protegido directamente (single-backend)
-Action Localize-Sandbox
-SandboxPath "C:\sandbox\Backend.accdb"
-SourceBackend "C:\proyecto\datos\Backend.accdb"
-Password "secret"
```

---

## Flujo de 3 etapas: sandbox → access-query → ERD

Esta es la narrativa completa que cubre los tres propósitos del skill.

### Preparación de entorno

```
PROYECTO      = C:\MiProyecto
BACKEND_ORIG  = C:\MiProyecto\datos\MiApp_Datos.accdb  (protegido con password)
FRONTEND      = C:\MiProyecto\MiApp.accdb
SANDBOX       = C:\Sandbox\MiApp_Datos.accdb
SIDECAR       = C:\MiProyecto\datos\MiApp_Datos_nopass.accdb
```

#### Etapa 1 — Crear y localize el sandbox

```powershell
# --- Paso 1: Copiar backend protegido a sandbox ---
& "C:\Users\adm\.config\opencode\.agent\skills\access-sandbox\SandboxManager.ps1" `
    -Action New-Sandbox `
    -SourceBackend "C:\MiProyecto\datos\MiApp_Datos.accdb" `
    -SandboxPath "C:\Sandbox\MiApp_Datos.accdb"

# --- Paso 2: Crear sidecar sin password (permite localize sin tocar el protegido) ---
& "C:\Users\adm\.config\opencode\.agent\skills\access-sandbox\SandboxManager.ps1" `
    -Action Make-Sidecar `
    -SourceBackend "C:\MiProyecto\datos\MiApp_Datos.accdb" `
    -Password "secret" `
    -SidecarSuffix "_nopass"

# --- Paso 3: Localizar sandbox (reemplazar vínculos con tablas locales) ---
& "C:\Users\adm\.config\opencode\.agent\skills\access-sandbox\SandboxManager.ps1" `
    -Action Localize-Sandbox `
    -SandboxPath "C:\Sandbox\MiApp_Datos.accdb" `
    -SourceSidecar "C:\MiProyecto\datos\MiApp_Datos_nopass.accdb"

# --- Verificación: confirmar que no quedan vínculos ---
$linked = & "C:\Users\adm\.config\opencode\.agent\skills\access-sandbox\SandboxManager.ps1" `
    -Action Discover-LinkedTables `
    -SandboxPath "C:\Sandbox\MiApp_Datos.accdb"

# Debería retornar array vacío
```

**Resultado:** `C:\Sandbox\MiApp_Datos.accdb` es una copia 100% autónoma del backend original,
sin vínculos externos, sin password.

---

#### Etapa 2 — Operar con access-query sobre el sandbox

```powershell
# Configurar password vacía para el sandbox (no tiene password)
$env:ACCESS_QUERY_PASSWORD = ""

# Consultar datos (lectura)
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" `
    -SQL "SELECT TOP 10 * FROM TbSolicitudes WHERE Estado = 'PENDIENTE'"

# Listar tablas y esquemas
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" -ListTables
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" -GetSchema -Table "TbSolicitudes"

# Seed de datos de test
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" `
    -Seed -AllowTable "TbSolicitudes" -FixtureTag "ERD_VERIFY" `
    -Exec "INSERT INTO TbSolicitudes (ID, Ref, Estado) VALUES (99901, 'TEST_ERD_01', 'PENDIENTE')"

# Cleanup después de verificar
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" `
    -Teardown -AllowTable "TbSolicitudes" -FixtureTag "ERD_VERIFY" `
    -Exec "DELETE FROM TbSolicitudes WHERE ID = 99901"
```

> `access-query` opera sobre **cualquier** backend (original o sandbox). Al apuntar al sandbox,
> todas las operaciones son seguras y no afectan al sistema real.

---

#### Etapa 3 — Generar ERD desde el sandbox

```powershell
# El sandbox ya es un ACCDB local, autocontenido, sin vínculos.
# Cualquier herramienta de generación de ERD puede apuntar directamente a él.

& ".\tools\Generar-ERD.ps1" `
    -BackendPath "C:\Sandbox\MiApp_Datos.accdb" `
    -Output "C:\Sandbox\ERD_MiApp.html"

# O con Access COM directamente (ejemplo conceptual):
$access = New-Object -ComObject Access.Application
$access.OpenCurrentDatabase("C:\Sandbox\MiApp_Datos.accdb", $false, "")
# ... usar DAO para extraer TableDefs y relationships ...
$access.CloseCurrentDatabase()
$access.Quit()
```

**Beneficio clave:** El ERD se genera desde datos reales (copiados al localize) y esquema completo,
pero sin tocar jamás el backend protegido original.

---

## Ejemplo CONDOR aplicado

CONDOR tiene **3 backends externos**: Expedientes_datos.accdb, NoConformidades_Datos.accdb y Lanzadera_Datos.accdb. El localize requiere un mapa de sidecars.

```powershell
# Rutas reales de CONDOR
$condorRoot = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR"
$localSandbox = "C:\sandboxes\CONDOR\condor_datos.accdb"
$sidecarRoot = "C:\sandboxes\CONDOR\sidecars"

# Mapa multi-backend: cada backend externo → su sidecar sin password
$backendMap = @{
    "$condorRoot\Expedientes_datos.accdb"      = "$sidecarRoot\Expedientes_datos_nopass.accdb"
    "$condorRoot\NoConformidades_Datos.accdb"   = "$sidecarRoot\NoConformidades_Datos_nopass.accdb"
    "$condorRoot\Lanzadera_Datos.accdb"         = "$sidecarRoot\Lanzadera_Datos_nopass.accdb"
    "$condorRoot\condor_datos.accdb"            = "$sidecarRoot\condor_datos_nopass.accdb"
} | ConvertTo-Json -Compress

# 1. Copiar backend principal a sandbox
& ".agent/skills/access-sandbox/SandboxManager.ps1" `
    -Action New-Sandbox `
    -SourceBackend "$condorRoot\condor_datos.accdb" `
    -SandboxPath $localSandbox

# 2. Crear sidecar sin password para c/ backend externo
foreach ($src in $backendMap.Keys) {
    $sidecar = $backendMap[$src]
    & ".agent/skills/access-sandbox/SandboxManager.ps1" `
        -Action Make-Sidecar `
        -SourceBackend $src `
        -Password "dpddpd" `
        -SidecarSuffix "_nopass"
}

# 3. Localizar sandbox (multi-backend map)
& ".agent/skills/access-sandbox/SandboxManager.ps1" `
    -Action Localize-Sandbox `
    -SandboxPath $localSandbox `
    -BackendSidecarMapJson $backendMap
```

---

## Parámetros comunes

| Parámetro | Default | Descripción |
|---|---|---|
| `-Action` | (requerido) | `New-Sandbox` \| `Discover-LinkedTables` \| `Make-Sidecar` \| `Localize-Sandbox` |
| `-SourceBackend` | | Ruta al backend fuente |
| `-SandboxPath` | | Ruta del sandbox destino |
| `-Password` | `""` | Password del backend protegido |
| `-SidecarSuffix` | `_nopass` | Sufijo para archivos sidecar |
| `-BackendSidecarMapJson` | | JSON hashtable `{SourceBackend -> SidecarPath}` para multi-backend |
| `-SourceSidecar` | | Ruta al sidecar (para localize sin password) |
| `-SandboxPassword` | `""` | Password del sandbox (si tiene) |
| `-WhatIf` | | Simula sin escribir |
| `-Verbose` | | Salida detallada |

---

## Notas de implementación

- **COM cleanup**: usa `FinalReleaseComObject` + `GC.Collect()` + `GC.WaitForPendingFinalizers()` para limpieza determinística
- **DAO versions**: intenta 160 → 150 → 140 → 120 en cascada
- **Compact**: `Localize-Sandbox` compacts la BD al final para limpiar espacio
- **Seguridad**: `AutomationSecurity = 1` desactiva prompts de macro
- **Headless**: `Visible = $false`, `UserControl = $false`
- **Unattended**: `SetWarnings($false)` se aplica al abrir la sesión — ningún prompt de Access aparece durante localize

---

## Flujo unattended (100% PowerShell, sin prompts)

El `Localize-Sandbox` es **completamente unattended** — ningún diálogo de Access, confirmación o input de usuario.

### Helpers canonicales (pattern de VBAManager.ps1)

```powershell
# Abrir session — configura Visible=$false, UserControl=$false,
#                 AutomationSecurity=1, SetWarnings=$false
$session = Open-AccessSession -AccessPath $SandboxPath -Password $SandboxPassword

# ... operaciones sobre la BD abierta ...

# Cerrar session — CloseCurrentDatabase + Quit + FinalReleaseComObject + GC
Close-AccessSession -Session $session
```

### Cómo ejecutar Localize-Sandbox sin prompts

```powershell
# Flujo completo sin password (sandboxes sin protección)
& "C:\Users\adm\.config\opencode\.agent\skills\access-sandbox\SandboxManager.ps1" `
    -Action Localize-Sandbox `
    -SandboxPath "C:\sandbox\MiApp_Datos.accdb" `
    -SourceSidecar "C:\proyecto\datos\MiApp_Datos_nopass.accdb"

# Con passwords — también unattended, solo pasar los parámetros
& "C:\Users\adm\.config\opencode\.agent\skills\access-sandbox\SandboxManager.ps1" `
    -Action Localize-Sandbox `
    -SandboxPath "C:\sandbox\MiApp_Datos.accdb" `
    -SourceBackend "C:\proyecto\datos\MiApp_Datos.accdb" `
    -Password "dpddpd" `
    -SandboxPassword "sandbox_secret"
```

> **Nota:** El path VBA/macro (usar `CurrentDb` desde macros AutoExec o módulos VBA para
> localize) está **DEPRECADO**. El flujo PowerShell con `Open-AccessSession` es el único
> camino soportado a partir de esta versión. No editar código VBA para localize.

---

## Relación con access-query

| Skill | access-sandbox | access-query |
|---|---|---|
| **Propósito** | Preparar un workspace local | Leer, consultar, modificar datos |
| **Operaciones** | New-Sandbox, Make-Sidecar, Localize-Sandbox | -SQL, -Exec, -Seed, -Teardown |
| **¿Modifica datos reales?** | No — solo crea archivos | Sí — pero en el target que se le pase |
| **Target natural** | Backend protegido → sandbox | Sandbox → consultas y fixtures |

**Flujo típico:**
1. `access-sandbox` crea y localize el sandbox
2. `access-query` opera sobre el sandbox (read/query/seed/teardown)
3. Una herramienta de ERD genera diagramas desde el sandbox

---

## Requisitos

- Windows con Microsoft Access instalado
- PowerShell 5.1+ o PowerShell 7+
- Permisos de lectura/escritura en las rutas especificadas
