# SKILL.md — Skill para workflow Access/VBA (Export → Trabajo → Sync → Compilar → ERD → Cierre)

## Objetivo
Definir un **skill** que automatice el workflow de desarrollo y documentación en un proyecto Microsoft Access/VBA:

1) **Al inicio** de una nueva feature/fix: **Exportar TODOS los módulos** del proyecto VBA a disco (snapshot base).
2) Se trabaja sobre la mejora editando los archivos exportados (normalmente con IA).
3) **Todo módulo modificado por la IA debe sincronizarse** (Import) hacia el VBA real de la BD.
4) Tras cada sincronización, el skill **debe proponer al usuario compilar** el proyecto en el VBE.
5) **Generación de documentación**: Extraer estructura de tablas (ERD/Diccionario) a Markdown para contexto de la IA.
6) **Al cerrar** la tarea (fin de sesión): export final opcional (snapshot consistente) + resumen.

> El skill es **autocontenido**: incluye `VBAManager.ps1` y todo lo necesario para ejecutarse sin dependencias externas.

---

## Alcance y supuestos
- Entorno: **Windows** con Microsoft Access instalado (automatización COM y DAO).
- El repositorio contiene una BD Access (`.accdb/.accde/.mdb/.mde`) en la raíz del proyecto, o el usuario la pasa por parámetro.
- La exportación se guarda bajo una carpeta configurable (default `src/`).
- La documentación se genera en `docs/ERD/` o ruta configurable.
- El skill cierra automáticamente cualquier instancia de Access que tenga abierta la BD objetivo (usando ROT para no afectar otras BDs abiertas).

---

## Capacidades del PS1 (VBAManager.ps1)

El PS1 es el motor que ejecuta todas las operaciones sobre Access. Acepta los siguientes parámetros:

| Parámetro | Tipo | Descripción |
|---|---|---|
| `-Action` | `Export\|Import\|Fix-Encoding\|Generate-ERD` | **Obligatorio**. Acción a ejecutar |
| `-AccessPath` | string | Ruta a la BD frontend |
| `-Password` | string | Contraseña de la BD (default en el script) |
| `-DestinationRoot` | string | Carpeta raíz de export/import (default: `src`) |
| `-ModuleName` | string[] | Uno o más nombres de módulo para operaciones selectivas |
| `-ImportMode` | `Auto\|Form\|Code` | Solo para `Import`: `Auto` (default), `Form` (forzar `.form.txt/.frm`), `Code` (forzar `.cls/.bas`) |
| `-BackendPath` | string | Ruta al backend `*_Datos.accdb` para Generate-ERD |
| `-ErdPath` | string | Carpeta de salida del ERD |
| `-Location` | `Both\|Src\|Access` | Ámbito de Fix-Encoding (default: `Both`) |

### Acciones

**`Export`** — Exporta módulos VBA de la BD a disco:
- Sin `-ModuleName`: exporta todos los módulos
- Con `-ModuleName A B C`: exporta solo los indicados
- Formularios: usa `SaveAsText` → `.form.txt` (UI + código) y `.cls` (solo código)
- Módulos/clases: usa `VBComponents.Export` → `.bas` / `.cls`

**`Import`** — Importa módulos desde disco a la BD:
- Sin `-ModuleName`: importa todos los archivos de `src/`
- Con `-ModuleName A B C`: importa solo los indicados
- Modo por defecto `-ImportMode Auto`: prioridad `.form.txt` > `.frm` > `.cls` > `.bas`
- `-ImportMode Form`: importa solo layout/formulario (`.form.txt/.frm`)
- `-ImportMode Code`: importa solo code-behind (`.cls/.bas`)
- Formularios (`.form.txt`): usa `LoadFromText` — completamente silencioso
- Módulos/clases: usa `DeleteLines + AddFromFile` — sin diálogos VBE

**`Fix-Encoding`** — Corrige encoding ANSI↔UTF-8:
- `-Location Src`: corrige BOM en archivos de `src/`
- `-Location Access`: reimporta desde `src/` para corregir en la BD
- `-Location Both` (default): ambos
- Selectivo con `-ModuleName`

**`Generate-ERD`** — Genera ERD en Markdown:
- Lee tablas, campos, tipos DAO, PKs e índices, y relaciones
- Detecta backends vinculados no alcanzables y los documenta
- Autodetecta `*_Datos.accdb` si no se pasa `-BackendPath`

---

## Comandos del CLI (cli.js)

El CLI es la interfaz de usuario sobre el PS1. Todos los comandos que expone mapean directamente a acciones del PS1.

### Tabla de comandos

| Comando CLI | Acción PS1 | Módulos | Descripción |
|---|---|---|---|
| `start` | `Export` | todos | Export inicial + inicia sesión |
| `watch` | `Import` (auto) | modificados | Auto-sync al detectar cambios en src/ |
| `export <Mod...>` | `Export` | selectivo | Exporta módulos específicos |
| `export-all` | `Export` | todos | Exporta todos sin iniciar sesión |
| `import <Mod...>` | `Import` | selectivo | Importa módulos específicos |
| `import-form <Mod...>` | `Import` (`ImportMode=Form`) | selectivo | Importa formularios desde `*.form.txt`/`*.frm` (UI + código) |
| `import-code <Mod...>` | `Import` (`ImportMode=Code`) | selectivo | Importa solo code-behind desde `*.cls`/`*.bas` |
| `import-all` | `Import` | todos | Importa todo src/ |
| `sync <Mod...>` | `Import` | selectivo | Alias de import |
| `fix-encoding [Mod...]` | `Fix-Encoding` | selectivo/todos | Corrige encoding |
| `generate-erd` | `Generate-ERD` | — | Genera ERD Markdown |
| `status` | — | — | Muestra estado de sesión |
| `end` | `Export` (opcional) | todos | Cierra sesión + export final |

### Ejecución secuencial obligatoria

⚠️ **Cada comando `import`, `import-code` o `import-form` cierra Access al inicio** (Kill via ROT). Por eso:

- **NUNCA encadenar comandos con `&&` o `;`** — el segundo comando falla porque Access ya fue cerrado por el primero
- **Ejecutar cada comando en su PROPIA llamada** — esperar el resultado antes de invocar el siguiente
- **Módulos normales (`import`) primero**, luego formularios (`import-code` o `import-form`) en una sola llamada si son varios

```powershell
# ✅ Correcto: uno por vez
node cli.js import NombreModulo --access "MiBD.accdb"
node cli.js import-code Form_A --access "MiBD.accdb"
node cli.js import-code Form_B Form_C --access "MiBD.accdb"

# ❌ Incorrecto: falla en el segundo comando
node cli.js import NombreModulo --access "MiBD.accdb" && node cli.js import-code Form_A --access "MiBD.accdb"
```

### Flags del CLI

| Flag | Mapea a PS1 | Default |
|---|---|---|
| `--access <ruta>` | `-AccessPath` | Autodetecta en CWD |
| `--password <pwd>` | `-Password` | — |
| `--destination_root <dir>` | `-DestinationRoot` | `src` |
| `--location Both\|Src\|Access` | `-Location` | `Both` |
| `--backend <ruta>` | `-BackendPath` | Autodetecta `*_Datos.accdb` |
| `--erd_path <dir>` | `-ErdPath` | `docs/ERD` |
| `--debounce_ms <n>` | — (Node.js) | `600` |
| `--auto_export_on_end false` | — (Node.js) | `true` |

---

## Estructura de archivos exportados

```
<DestinationRoot>/
├── modules/          # Tipo 1 (vbext_ct_StdModule)   → .bas
├── classes/          # Tipo 2 (vbext_ct_ClassModule)  → .cls
└── forms/            # Tipo 3/100 (formularios)
    ├── Form_X.form.txt   # UI + código (SaveAsText, acForm=2)
    └── Form_X.cls        # Solo código VBA
```

Ejemplo real:
```
src/modules/VariablesGlobales.bas
src/classes/CUsuario.cls
src/forms/Form_FormGestion.form.txt
src/forms/Form_FormGestion.cls
```

---

## Comportamiento de formularios

Los formularios Access tienen tratamiento especial respecto a módulos y clases:

- **Export**: `Application.SaveAsText(2, nombreSinPrefixForm_, ruta)`. El nombre del objeto Access es el nombre del VBComponent **sin** el prefijo `Form_`.
- **Import**: `Application.LoadFromText(2, nombreSinPrefixForm_, ruta)`. Nunca usa `VBComponents.Import()`.
- **Fallback en export**: si `SaveAsText` falla, usa `component.Export()` y registra el aviso.
- Los formularios también generan un `.cls` paralelo con solo el código VBA para facilitar diff.

---

## Regla de oro: código en .cls, UI en .form.txt, sync automático

**Nunca editar el CodeBehind del `.form.txt` directamente.** El flujo correcto es:

1. **CAMBIO DE CÓDIGO VBA** → editar **SOLO el `.cls`**
2. **CAMBIO DE UI** (propiedades de controles, layout) → editar **SOLO el `.form.txt`**
3. **Antes de importar** (modo Code o Auto) → el handler sincroniza automáticamente el CodeBehind del `.form.txt` con el contenido del `.cls`

### Por qué

El `.cls` es el archivo "maestro" para código VBA. El `.form.txt` tiene dos secciones:
- **UI** (antes de `CodeBehind`): propiedades de controles, layout
- **CodeBehind** (después de `CodeBehind`): código VBA — **es un espejo del `.cls`**

Cuando Access exporta un formulario, guarda el código en ambas secciones por separado. Si la IA modifica solo el `.cls` pero el CodeBehind del `.form.txt` está desincronizado, los cambios no se aplican correctamente.

### Cómo funciona el sync automático

En `import-modules` (handler.js), antes de invocar VBAManager.ps1 con `-ImportMode Code` o `-ImportMode Auto`:
1. Se lee el contenido del `.cls`
2. Se reemplaza la sección `CodeBehind` del `.form.txt` con ese contenido
3. Se importa el `.form.txt` ya sincronizado a Access

### Verificación manual

Para verificar que `.cls` y `CodeBehind` coinciden sin hacer import:

```powershell
# Mostrar diferencias entre .cls y CodeBehind del .form.txt
node cli.js verify-code Form_subfrmDatosCDCA_Generales
```

(El comando `verify-code` aún no existe — por ahora verificar manualmente con diff o hacer import-code y confirmar que los cambios aparecen en Access.)

---

## Flujo de modificación de UI de formularios por la IA

Este es un caso de uso de primera clase del skill: **la IA modifica la interfaz de un formulario editando su `.form.txt` y el cambio se aplica a la BD**.

### El ciclo completo

```
export Form_X  →  IA edita src/forms/Form_X.form.txt  →  import Form_X
```

El `.form.txt` contiene tanto la definición de controles y propiedades (UI) como el código VBA. Al importarlo con `LoadFromText`, Access recarga el formulario completo — UI y código — reflejando todos los cambios.

### Comandos para este flujo

```powershell
# 1. Exportar el formulario a src/ (si no está ya exportado)
node cli.js export Form_FormGestion

# 2. La IA edita src/forms/Form_FormGestion.form.txt
#    (propiedades de controles, nuevos botones, layout, etc.)

# 3. Importar el formulario modificado a la BD
node cli.js import Form_FormGestion

# Varios formularios a la vez:
node cli.js import Form_FormGestion Form_subfrmDetalle Form_FormBusqueda
```

### En modo watch (automático)

Con `watch` activo, cada vez que la IA guarda un `.form.txt` se importa automáticamente a la BD sin ninguna intervención:

```powershell
node cli.js watch
# A partir de aquí: la IA edita y guarda .form.txt → se importa solo
```

### Regla crítica: nunca crear un `.form.txt` desde cero

La IA **debe partir siempre de un `.form.txt` exportado con `export`**, nunca generarlo de cero. El formato es propietario de Access — si se introduce una línea mal formada, `LoadFromText` falla. El flujo correcto es siempre exportar primero para tener la base correcta y luego modificar.

### Qué puede modificar la IA en el `.form.txt`

El archivo tiene esta estructura general:

```
Version =21
VersionRequired =20
Begin Form
    Width = ...
    Caption = "..."
    Begin Label etiqueta1
        Left = ...
        Top = ...
        Caption = "Nombre:"
    End
    Begin TextBox txtNombre
        Left = ...
        Width = ...
        ControlSource = "NombreCampo"
    End
    ...
End
CodeBehind
    ' Aquí va el código VBA del formulario
```

La IA puede modificar con seguridad:
- Propiedades de controles existentes (`Left`, `Top`, `Width`, `Height`, `Caption`, `Visible`, etc.)
- Valores de `ControlSource`, `RowSource`, `DefaultValue`
- Añadir nuevos controles copiando la estructura de uno existente

Debe evitar:
- Editar manualmente la sección `CodeBehind` del `.form.txt` (el skill la sincroniza desde el `.cls`)
- Cambiar el `GUID` o `Checksum` del formulario (Access los regenera en el siguiente export)
- Renombrar controles que tengan referencias en el código VBA sin actualizar también el código
- Alterar la estructura de bloques `Begin`/`End` de forma que queden mal anidados

---

## Comportamiento de sesión

El skill persiste el estado en `.access-vba-skill/session.json` en la raíz del proyecto:

```json
{
  "active": true,
  "startedAt": "2025-01-15T09:00:00.000Z",
  "accessPath": "C:\\proyecto\\MiBD.accdb",
  "destinationRoot": "C:\\proyecto\\src",
  "modulesPath": "C:\\proyecto\\src",
  "changedModules": ["Utilidades", "Form_FormGestion"],
  "lastSyncAt": "2025-01-15T09:45:00.000Z",
  "pendingModules": [],
  "watcherPid": null
}
```

Esto permite que `import`, `sync`, `end` y `status` funcionen sin mantener un proceso largo vivo.

---

## Gestión de Access (headless)

Antes de abrir la BD el PS1 deshabilita temporalmente mediante DAO:
- **AllowBypassKey**: habilitado para acceso sin restricciones
- **StartupForm**: eliminado para que no abra el formulario de inicio
- **AutoExec**: renombrado a `AutoExec_TraeBackup` para evitar ejecución

La BD se abre con:
```powershell
$access.Visible = $false
$access.UserControl = $false
$access.AutomationSecurity = 1   # desactiva seguridad de macros
$access.DoCmd.SetWarnings($false) # suprime diálogos de Access
```

Al cerrar, todos los valores se restauran al estado original.

---

## Estructura del paquete

```
<projectRoot>/
  access-vba-sync/
    VBAManager.ps1     # Motor PowerShell (Export/Import/Fix-Encoding/Generate-ERD)
    handler.js         # Lógica Node.js (sesión, debounce, watcher, métodos del skill)
    cli.js             # Interfaz de comandos
    README.md          # Documentación de uso
    SKILL.md           # Este documento (especificación)
```

> El skill vive en su carpeta pero opera con `projectRoot = cwd`, de forma que `src/` y `docs/ERD/` quedan en el proyecto y no dentro del skill.

---

## Flujo de trabajo esperado

### Nueva feature/fix
```
start → (generate-erd) → IA edita src/ → watch|import → compilar VBE → end
```

1. `start` — Export total, snapshot base
2. `generate-erd` — Contexto de datos para la IA (opcional)
3. IA modifica archivos en `src/`
4. `watch` (automático) o `import <módulos>` (manual)
5. Compilar: **Abre Access → VBE → Debug → Compile**
6. `end` — Sync final + export final opcional + resumen

---

## Casos límite cubiertos

- Varias BDs en CWD → elección determinista (alfabético) + warning
- Ruta relativa en `--access` → se resuelve contra CWD
- Nombre de módulo sin extensión → `Resolve-ImportFileForModule` busca por prioridad: `.form.txt` > `.frm` > `.cls` > `.bas`
- Formularios sin prefijo `Form_` → detectados por tipo de componente (tipo 3 o 100)
- Guardados masivos simultáneos → batching con debounce configurable
- Access abierto/bloqueado → error claro, no loops infinitos
- BD con contraseña → `--password` propagado a todas las operaciones DAO y COM

---

## Pruebas mínimas

- `start` con BD única → crea `src/modules/`, `src/classes/`, `src/forms/`
- `start` sin BD en CWD → error claro
- `export Form_X` → genera `src/forms/Form_X.form.txt` y `src/forms/Form_X.cls`
- `import Form_X` → reimporta el formulario sin diálogos
- `watch`: editar un `.bas` → se importa automáticamente tras debounce
- `watch`: editar un `.form.txt` → se importa automáticamente
- `fix-encoding --location Src` → solo toca archivos en `src/`
- `import-all` → importa todos los archivos de `src/` sin intervención
- `generate-erd` → genera `docs/ERD/NombreBackend.md`
- `end` → export final + resumen con módulos tocados
