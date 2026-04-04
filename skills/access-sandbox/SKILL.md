# SKILL.md — access-sandbox: Provisioning de Workspaces Access Locales

## Objetivo

`access-sandbox` prepara un **workspace Access local y autocontenido** a partir de un backend fuente (`.accdb`/`.mdb`). El resultado es un entorno de trabajo aislado, sin vínculos externos, que puede usarse para:

1. **Generación de ERDs** — crear diagramas de estructura de datos sin tocar backends remotos o protegidos
2. **Consultas y operaciones de datos** — alimentar a `access-query` con un objetivo seguro
3. **Testing** — ejecutar tests sobre una copia exacta que no afecta al sistema real

> **Nota:** Esta skill **prepara** el sandbox. Para leer, consultar, modificar o limpiar datos en ese sandbox, usar `access-query` (que es la skill complementaria).

---

## División de responsabilidades

| Skill | Rol | ¿Modifica datos? |
|---|---|---|
| `access-sandbox` | Provisioning: copia, sidecar, localize | Solo crea archivos |
| `access-query` | Read/query/seed/cleanup sobre un backend existente | Sí (INSERT/UPDATE/DELETE con guardas) |

---

## Supuestos

- **Windows** con Microsoft Access instalado (COM + DAO)
- **PowerShell 5.1+** (o PowerShell 7+)
- La BD fuente (backend) debe estar **cerrada** antes de invocar el script
- El script abre Access en modo **headless** (invisible, sin UserControl)
- Acceso a `DAO.DBEngine` (versiones 120/140/150/160 compatibles)

---

## Capacidades del PS1 (SandboxManager.ps1)

### Parámetros

| Parámetro | Tipo | Requerido | Descripción |
|---|---|---|---|
| `-Action` | `string` | Sí | `New-Sandbox` \| `Discover-LinkedTables` \| `Make-Sidecar` \| `Localize-Sandbox` |
| `-SourceBackend` | `string` | Sí* | Ruta absoluta al backend fuente (`.accdb`/`.mdb`) |
| `-SandboxPath` | `string` | Sí* | Ruta absoluta de destino del sandbox (copia del backend) |
| `-Password` | `string` | No | Contraseña del backend protegido (si corresponde) |
| `-SidecarSuffix` | `string` | No | Sufijo para el sidecar sin password (default: `_nopass`) |
| `-BackendSidecarMapJson` | `string` | No* | JSON string de hashtable `{SourceBackend -> SidecarPath}` para proyectos con múltiples backends externos |
| `-SourceSidecar` | `string` | No | Ruta a un sidecar existente para usar como fuente de datos (evita abrir el protegido) |
| `-SandboxPassword` | `string` | No | Contraseña para el sandbox si se quiere proteger (default: sin password) |
| `-WhatIf` | `switch` | No | Simula sin escribir archivos |
| `-Verbose` | `switch` | No | Salida detallada |

*`Localize-Sandbox` requiere `-BackendSidecarMapJson` cuando el sandbox tiene tablas vinculadas a **múltiples** backends externos (ej: CONDOR con Expedientes_datos, NoConformidades_datos, Lanzadera_datos). Si solo hay un backend, puede usarse `-SourceSidecar` directamente.
*`New-Sandbox` y `Localize-Sandbox` requieren `-SourceBackend` y `-SandboxPath`.
*`Discover-LinkedTables` requiere `-SandboxPath` únicamente.

---

## Acciones

### `New-Sandbox`

Crea una copia exacta del backend fuente en la ruta del sandbox:

```
SandboxManager.ps1 -Action New-Sandbox -SourceBackend "C:\proyecto\datos\MiBackend.accdb" -SandboxPath "C:\sandbox\MiBackend.accdb"
```

- Valida que `-SourceBackend` exista
- Si `-SandboxPath` termina en `.accdb`/`.mdb` sin carpeta, crea la carpeta destino
- Sobrescribe el sandbox si ya existe (confirmación implícita en flujo desatenido)
- **No abre Access** — usa `Copy-Item` directo (más rápido y seguro)

### `Discover-LinkedTables`

Abre el sandbox (headless) y descubre todas las tablas vinculadas:

```
SandboxManager.ps1 -Action Discover-LinkedTables -SandboxPath "C:\sandbox\MiBackend.accdb" -Password ""
```

Retorna por pipeline un objeto:

```powershell
[PSCustomObject]@{
    TableName        = "tblContratos"
    Connect          = ";DATABASE=C:\proyecto\datos\MiBackend.accdb"
    SourceTable      = "tblContratos"
    IsLinked         = $true
}
```

### `Make-Sidecar`

Crea una copia sin password de un backend protegido:

```
SandboxManager.ps1 -Action Make-Sidecar -SourceBackend "C:\proyecto\datos\MiBackend.accdb" -Password "secret" -SidecarSuffix "_nopass"
```

Genera: `C:\proyecto\datos\MiBackend_nopass.accdb`

- Usa `DAO.DBEngine.OpenDatabase` con `;PWD=` para abrir el protegido
- Crea un nuevo archivo ACCDB vacío
- Itera todas las `TableDefs` del protegido y las copia al sidecar:
  - Tablas locales: copy completa (datos + esquema)
  - Tablas vinculadas: se recrean como **tablas locales** en el sidecar con los mismos datos de origen
- El sidecar queda **sin contraseña**
- Útil cuando el sandbox necesita acceder a datos de un backend protegido sin conocer la password

### `Localize-Sandbox`

Abre el sandbox, detecta tablas vinculadas, y las convierte en tablas locales usando datos del source (sidecar o backends múltiples):

**Single-backend (un solo backend externo):**
```
SandboxManager.ps1 -Action Localize-Sandbox -SandboxPath "C:\sandbox\MiBackend.accdb" -SourceSidecar "C:\proyecto\datos\MiBackend_nopass.accdb"
```

**Multi-backend (múltiples backends externos — requiere `-BackendSidecarMapJson`):**
```
$map = @{
    "C:\proyecto\datos\Backend1.accdb" = "C:\proyecto\datos\Backend1_nopass.accdb"
    "C:\proyecto\datos\Backend2.accdb" = "C:\proyecto\datos\Backend2_nopass.accdb"
} | ConvertTo-Json -Compress

SandboxManager.ps1 -Action Localize-Sandbox -SandboxPath "C:\sandbox\MiBackend.accdb" -BackendSidecarMapJson $map
```

Flujo interno:
1. `Discover-LinkedTables` → obtiene lista de tablas vinculadas
2. Para cada tabla vinculada, usa DAO para:
   - Resolver el backend source de la tabla (desde `TableDef.Connect`)
   - Obtener el sidecar correspondiente del mapa (multi-backend) o usar `-SourceSidecar` (single-backend)
   - Eliminar la `TableDef` vinculada del sandbox
   - Crear una nueva `TableDef` local con el mismo esquema
   - Copiar todos los registros desde el source sidecar
3. El sandbox queda **100% autónomo** — sin vínculos externos

---

## Arquitectura COM

El script sigue el patrón canónico de `access-vba-sync/VBAManager.ps1`:

```powershell
# helpers canonicales (disponibles en SandboxManager.ps1)
# Open-AccessSession — abre session unattended con tracking de PID
$session = Open-AccessSession -AccessPath $AccessPath -Password $Password
$access  = $session.AccessApplication
$accessPid = $session.ProcessId

# Acceso a BD abierto — sin prompts (SetWarnings=$false ya aplicado)
$access.DoCmd.SetWarnings($false)

# ... operaciones ...

# Cierre limpio con Close-AccessSession
Close-AccessSession -Session $session
```

> **Phase 2 (VBA-based):** El flujo que usaba VBA/macros como paso intermedio para
> localize está **DEPRECADO**. El camino correcto es el flujo PowerShell 100% unattended
> descrito en este documento. No se debe usar `DoCmd.DeleteObject` ni `CurrentDb` desde
> macros VBA para localize — esa ruta ya no es el path recomendado.

Para DAO:

```powershell
$engine = New-Object -ComObject DAO.DBEngine.120  # o 140/150/160
$db = $engine.OpenDatabase($Path, $false, $false, $ConnectionString)
$db.Close()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($db) | Out-Null
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
```

---

## Estructura del skill

```
.agent/skills/access-sandbox/
├── SKILL.md              # Este archivo (especificación)
├── SandboxManager.ps1    # Motor PowerShell
├── README.md             # Guía de uso + ejemplos
```

> El skill vive en su carpeta pero todas las rutas de operación son **absolutas y parametrizadas**, de forma que opera sobre cualquier proyecto.

---

## Flujo de 3 etapas (narrativa completa)

### Etapa 1 — Crear / localizar el sandbox

```
PROYECTO      = C:\MiProyecto
BACKEND_ORIG  = C:\MiProyecto\datos\MiApp_Datos.accdb  (protegido)
SANDBOX       = C:\Sandbox\MiApp_Datos.accdb
SIDECAR       = C:\MiProyecto\datos\MiApp_Datos_nopass.accdb
```

```powershell
# 1. Copiar backend protegido a sandbox
& ".agent/skills/access-sandbox/SandboxManager.ps1" `
    -Action New-Sandbox `
    -SourceBackend "C:\MiProyecto\datos\MiApp_Datos.accdb" `
    -SandboxPath "C:\Sandbox\MiApp_Datos.accdb"

# 2. Crear sidecar sin password (permite localize sin exponer el protected)
& ".agent/skills/access-sandbox/SandboxManager.ps1" `
    -Action Make-Sidecar `
    -SourceBackend "C:\MiProyecto\datos\MiApp_Datos.accdb" `
    -Password "secret" `
    -SidecarSuffix "_nopass"

# 3. Localizar sandbox (reemplazar vínculos con tablas locales)
& ".agent/skills/access-sandbox/SandboxManager.ps1" `
    -Action Localize-Sandbox `
    -SandboxPath "C:\Sandbox\MiApp_Datos.accdb" `
    -SourceSidecar "C:\MiProyecto\datos\MiApp_Datos_nopass.accdb"
```

**Resultado:** `C:\Sandbox\MiApp_Datos.accdb` es ahora un backend local, autocontenido, sin password, sin vínculos externos.

---

### Etapa 2 — Operar sobre el sandbox con access-query

Una vez que el sandbox está localizeado, `access-query` puede operar sobre él sin riesgo de tocar el original:

```powershell
# Consultar datos (lectura - sin restricciones de seguridad extra)
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" `
    -SQL "SELECT TOP 10 * FROM TbSolicitudes"

# Listar tablas
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" -ListTables

# Seed de datos de prueba
$env:ACCESS_QUERY_PASSWORD = ""  # sandbox no tiene password
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" `
    -Seed -AllowTable "TbSolicitudes" -FixtureTag "ERD_TEST" `
    -Exec "INSERT INTO TbSolicitudes (ID, Ref) VALUES (99901, 'TEST_ERD_01')"

# Cleanup después de tests
.\query-backend.ps1 -BackendPath "C:\Sandbox\MiApp_Datos.accdb" `
    -Teardown -AllowTable "TbSolicitudes" -FixtureTag "ERD_TEST" `
    -Exec "DELETE FROM TbSolicitudes WHERE ID BETWEEN 99901 AND 99910"
```

> **Nota:** `access-query` tiene su propia gestión de passwords. Si el sandbox no tiene password, usar `""` o la env var `ACCESS_QUERY_PASSWORD` vacía.

---

### Etapa 3 — Generar ERD desde el sandbox

El sandbox localizeado es un objetivo seguro para generación de ERDs (no toca el backend protegido):

```
# Generar ERD desde el sandbox (herramienta de tu elección, ejemplo conceptual)
# El sandbox en C:\Sandbox\MiApp_Datos.accdb ya tiene:
#   - Todas las tablas como locales
#   - Datos copiados (para referencias完整性)
#   - Sin password
#   - Sin vínculos externos

& ".\tools\Generar-ERD.ps1" -BackendPath "C:\Sandbox\MiApp_Datos.accdb" -Output "C:\Sandbox\ERD_MiApp.html"
```

El resultado es un ERD del esquema completo sin haber tocado `MiApp_Datos.accdb` original.

---

## Casos de uso cubiertos

| Caso | Etapa 1 | Etapa 2 | Etapa 3 |
|---|---|---|---|
| Generar ERD sin tocar origen | `Localize-Sandbox` | (opcional) `access-query` para verificar | `Generar-ERD` → sandbox |
| Explorar datos protegidos con `access-query` | `Make-Sidecar` + `Localize-Sandbox` | `access-query` → sandbox | N/A |
| Testing con fixtures reales | `New-Sandbox` + `Localize-Sandbox` | `access-query -Seed` → sandbox | Tests automatizados |
| Comparar esquemas (dev vs prod) | Dos `Localize-Sandbox` | `access-query -Compare` | Análisis diff |
| Seed masivo sin riesgo | `Localize-Sandbox` | `access-query -Seed -AllowTable` → sandbox | Verificación |

---

## Casos límite

- **Backend protegido sin sidecar** → `Localize-Sandbox` necesita `-SourceSidecar` o fallará
- **Sandbox ya localizeado** → `Discover-LinkedTables` retornará vacío (sin vínculos)
- **Backend fuente bloqueado por otro proceso** → error claro con path del proceso bloqueante
- **ACCESS NO instalado** → el script valida `New-Object -ComObject Access.Application` al inicio y muere con error legible
- **Versiones DAO no disponibles** → intenta 160→150→140→120→36 en cascada

---

## Pruebas mínimas

- `New-Sandbox` → archivo copiado, tamaño > 0, tamaño igual al fuente (copia bit-a-bit)
- `Discover-LinkedTables` sobre sandbox sin vínculos → array vacío
- `Discover-LinkedTables` sobre sandbox con vínculos → array con n elementos `{TableName, Connect, SourceTable}`
- `Make-Sidecar` → archivo generado, `OpenDatabase` sin password exitoso
- `Localize-Sandbox` → después de la operación, `Discover-LinkedTables` retorna array vacío
