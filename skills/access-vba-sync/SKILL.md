# SKILL.md — Skill para workflow Access/VBA (Export → Trabajo → Sync → Compilar → ERD → Cierre)

## Objetivo
Definir un **skill** que automatice el workflow de desarrollo y documentación en un proyecto Microsoft Access/VBA:

1) **Al inicio** de una nueva feature/fix: **Exportar TODOS los módulos** del proyecto VBA a disco (snapshot base).
2) Se trabaja sobre la mejora editando los archivos exportados (normalmente con IA).
3) **Todo módulo modificado por la IA debe sincronizarse** (Import) hacia el VBA real de la BD.
4) Tras cada sincronización, el skill **debe proponer al usuario compilar** el proyecto en el VBE.
5) **Generación de documentación**: Extraer estructura de tablas (ERD/Diccionario) a Markdown para contexto de la IA.
6) **Gestión de módulos**: Borrar, renombrar y listar módulos directamente desde el CLI.
7) **Al cerrar** la tarea (fin de sesión): export final opcional (snapshot consistente) + resumen.

> El skill es **autocontenido**: incluye `VBAManager.ps1` y todo lo necesario para ejecutarse sin dependencias externas.

---

## Alcance y supuestos
- Entorno: **Windows** con Microsoft Access instalado (automatización COM y DAO).
- El repositorio contiene una BD Access (`.accdb/.accde/.mdb/.mde`) en la raíz del proyecto, o el usuario la pasa por parámetro con `--access`.
- `--access` acepta rutas absolutas o relativas — la BD **no necesita estar en CWD**.
- La exportación se guarda bajo una carpeta configurable (default `src/`), configurable con `--destination_root`.
- La documentación se genera en `docs/ERD/` o ruta configurable.
- **La BD debe estar cerrada** antes de ejecutar cualquier comando. El skill abre Access en modo headless.

---

## Capacidades del PS1 (VBAManager.ps1)

El PS1 es el motor que ejecuta todas las operaciones sobre Access. Acepta los siguientes parámetros:

| Parámetro | Tipo | Descripción |
|---|---|---|
| `-Action` | `Export\|Import\|Fix-Encoding\|Generate-ERD\|Delete\|Rename\|List` | **Obligatorio**. Acción a ejecutar |
| `-AccessPath` | string | Ruta a la BD frontend |
| `-Password` | string | Contraseña de la BD (default en el script) |
| `-DestinationRoot` | string | Carpeta raíz de export/import (default: `src`) |
| `-ModuleName` | string[] | Uno o más nombres de módulo para operaciones selectivas |
| `-ImportMode` | `Auto\|Form\|Code` | Solo para `Import`: `Auto` (default), `Form` (forzar `.form.txt/.frm`), `Code` (forzar `.cls/.bas`) |
| `-NewModuleName` | string | Solo para `Rename`: el nuevo nombre del módulo |
| `-DeleteFromSrc` | switch | Solo para `Delete`: si se pasa, también borra los archivos de `src/` |
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

**`Delete`** — Borra módulos del proyecto VBA:
- Requiere `-ModuleName` con al menos un módulo
- Formularios: usa `DoCmd.DeleteObject(acForm, nombre)` (el nombre de Access, sin prefijo `Form_`)
- Módulos/clases: usa `VBComponents.Remove`
- Con `-DeleteFromSrc`: también borra los archivos correspondientes (`.form.txt`, `.cls`, `.bas`) de `src/`
- Sin `-DeleteFromSrc`: solo borra de la BD, los archivos de `src/` permanecen intactos

**`Rename`** — Renombra un módulo en la BD y en `src/`:
- Requiere exactamente un `-ModuleName` (el nombre actual) y `-NewModuleName` (el nuevo nombre)
- Formularios: usa `DoCmd.Rename(nuevoNombre, acForm, nombreActual)`
- Módulos/clases: asigna directamente `component.Name = nuevoNombre`
- Automáticamente renombra los archivos correspondientes en `src/` (`.form.txt`, `.cls`, `.bas`)
- Si no encuentra archivos en `src/`, sugiere ejecutar `export <NuevoNombre>` para generarlos

**`List`** — Lista todos los módulos VBA del proyecto:
- No requiere `-ModuleName`
- Muestra: nombre, tipo (Module/Class/Form), número de líneas de código
- La salida se emite por stdout con colores por tipo

**`Fix-Encoding`** — Corrige encoding ANSI↔UTF-8:
- `-Location Src`: corrige BOM en archivos de `src/`
- `-Location Access`: reimporta desde `src/` para corregir en la BD
- `-Location Both` (default): ambos
- Selectivo con `-ModuleName`

**`Generate-ERD`** — Genera ERD en Markdown:
- Lee tablas, campos, tipos DAO, PKs e índices, y relaciones
- Detecta tablas vinculadas usando `TableDef.Attributes` (bits `dbAttachedTable`, `dbAttachedODBC`) y `TableDef.Connect`
- Clasifica vinculaciones por tipo: Access, ODBC, Excel, SharePoint, Text/CSV, HTML, Otro
- Extrae `SourceTableName` cuando difiere del nombre local
- Comprueba alcanzabilidad de orígenes basados en fichero (Access, Excel, Text)
- Documenta conexiones ODBC con DSN/Driver/Server sin intentar comprobar alcanzabilidad
- Acepta `-AccessPath` para generar ERD del frontend (mostrando tablas vinculadas)
- Acepta `-BackendPath` para generar ERD del backend (tablas locales)
- Si se pasan ambos, genera dos ficheros `.md`
- Si no se pasa ninguno, auto-detecta `*_Datos.accdb`; si no encuentra, usa el frontend

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
| `delete <Mod...>` | `Delete` | selectivo | Borra módulos de la BD (y de src/ con `--delete-src`) |
| `rename <Old> <New>` | `Rename` | uno | Renombra un módulo en la BD y en src/ |
| `list` | `List` | — | Lista todos los módulos VBA de la BD |
| `fix-encoding [Mod...]` | `Fix-Encoding` | selectivo/todos | Corrige encoding |
| `generate-erd` | `Generate-ERD` | — | Genera ERD Markdown (frontend, backend, o ambos) |
| `status` | — | — | Muestra estado de sesión |
| `end` | `Export` (opcional) | todos | Cierra sesión + export final |

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
| `--delete-src` | `-DeleteFromSrc` | `false` |

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

La carpeta destino se controla con `--destination_root`:
```powershell
# Exportar a una carpeta distinta de src/
node cli.js start --destination_root "D:\exports\mi_proyecto"

# Todos los comandos respetan la carpeta configurada
node cli.js import Utilidades --destination_root "D:\exports\mi_proyecto"
```

---

## Comportamiento de formularios

Los formularios Access tienen tratamiento especial respecto a módulos y clases:

- **Export**: `Application.SaveAsText(2, nombreSinPrefixForm_, ruta)`. El nombre del objeto Access es el nombre del VBComponent **sin** el prefijo `Form_`.
- **Import**: `Application.LoadFromText(2, nombreSinPrefixForm_, ruta)`. Nunca usa `VBComponents.Import()`.
- **Delete**: `DoCmd.DeleteObject(2, nombreSinPrefixForm_)`. No usa `VBComponents.Remove()`.
- **Rename**: `DoCmd.Rename(nuevoNombre, 2, nombreActual)`. No usa `component.Name =`.
- **Fallback en export**: si `SaveAsText` falla, usa `component.Export()` y registra el aviso.
- Los formularios también generan un `.cls` paralelo con solo el código VBA para facilitar diff.

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
- Modificar el código VBA en la sección `CodeBehind`

Debe evitar:
- Cambiar el `GUID` o `Checksum` del formulario (Access los regenera en el siguiente export)
- Renombrar controles que tengan referencias en el código VBA sin actualizar también el código
- Alterar la estructura de bloques `Begin`/`End` de forma que queden mal anidados

---

## Gestión de módulos: Delete, Rename, List

### Borrar módulos

```powershell
# Borrar de la BD solamente (archivos de src/ permanecen)
node cli.js delete modObsoleto Form_FormViejo

# Borrar de la BD Y de src/
node cli.js delete modObsoleto Form_FormViejo --delete-src
```

El delete opera sobre módulos, clases y formularios. Para formularios usa `DoCmd.DeleteObject` (nunca `VBComponents.Remove`, que no funciona con forms). Cada módulo se procesa individualmente y los errores se reportan sin detener el lote.

### Renombrar módulos

```powershell
# Renombrar un módulo (BD + archivos src/ automáticamente)
node cli.js rename modNombreViejo modNombreNuevo

# Renombrar un formulario
node cli.js rename Form_FormViejo Form_FormNuevo
```

Solo se puede renombrar un módulo a la vez. El rename actualiza tanto la BD como los archivos en `src/` (`.form.txt`, `.cls`, `.bas`). Si no encuentra archivos en `src/`, sugiere un export del nuevo nombre.

**Advertencia**: renombrar un módulo no actualiza automáticamente las referencias a ese módulo en otros módulos. La IA o el usuario debe buscar y reemplazar `modNombreViejo` → `modNombreNuevo` en el código que lo referencia.

### Listar módulos

```powershell
node cli.js list
```

Muestra una tabla con todos los módulos VBA del proyecto: nombre, tipo (Module/Class/Form) y número de líneas de código. Los tipos se muestran con colores distintos. Útil para que la IA conozca la estructura del proyecto antes de trabajar.

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

Los comandos `delete`, `rename`, `list`, `export` y `export-all` **no requieren sesión activa** — detectan la BD automáticamente o la reciben con `--access`.

---

## Gestión de Access (headless)

Antes de abrir la BD el PS1 deshabilita temporalmente mediante DAO:
- **AllowBypassKey**: habilitado para acceso sin restricciones
- **StartupForm**: eliminado para que no abra el formulario de inicio
- **AutoExec**: renombrado a `AutoExec_TraeBackup` para evitar su ejecución

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
    VBAManager.ps1     # Motor PowerShell (Export/Import/Delete/Rename/List/Fix-Encoding/Generate-ERD)
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
start → (generate-erd) → (list) → IA edita src/ → watch|import → compilar VBE → end
```

1. `start` — Export total, snapshot base
2. `generate-erd` — Contexto de datos para la IA (backend)
3. `generate-erd --access MiFrontend.accdb` — Mapa de tablas vinculadas del frontend (opcional)
4. `list` — Ver qué módulos existen (opcional, útil para la IA)
5. IA modifica archivos en `src/`
6. `watch` (automático) o `import <módulos>` (manual)
7. Compilar: **Abre Access → VBE → Debug → Compile**
8. `end` — Sync final + export final opcional + resumen

### Refactoring / limpieza
```
list → delete modObsoleto --delete-src → rename modViejo modNuevo → export modNuevo
```

### Uso desde fuera de CWD
```powershell
node cli.js list --access "D:\proyectos\MiBD.accdb"
node cli.js export-all --access "D:\proyectos\MiBD.accdb" --destination_root "D:\exports\src"
```

---

## Casos límite cubiertos

- Varias BDs en CWD → elección determinista (alfabético) + warning
- `--access` con ruta absoluta fuera de CWD → funciona correctamente
- Ruta relativa en `--access` → se resuelve contra CWD
- `--destination_root` con ruta absoluta → los módulos van al directorio indicado
- Nombre de módulo sin extensión → `Resolve-ImportFileForModule` busca por prioridad: `.form.txt` > `.frm` > `.cls` > `.bas`
- Formularios sin prefijo `Form_` → detectados por tipo de componente (tipo 3 o 100)
- Guardados masivos simultáneos → batching con debounce configurable
- Access abierto/bloqueado → error claro, no loops infinitos
- BD con contraseña → `--password` propagado a todas las operaciones DAO y COM
- `delete` de formularios → usa `DoCmd.DeleteObject` (no `VBComponents.Remove`)
- `rename` de formularios → usa `DoCmd.Rename` (no `component.Name =`)
- `rename` sin archivos en `src/` → avisa y sugiere export
- `delete` con `--delete-src` → borra `.form.txt`, `.cls` y `.bas` del módulo en todas las subcarpetas
- `export` selectivo sin sesión activa → no hace un export-all previo innecesario

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
- `generate-erd` → auto-detecta backend, genera `docs/ERD/NombreBackend.md`
- `generate-erd --access MiFrontend.accdb` → genera ERD del frontend con tablas vinculadas clasificadas por tipo
- `generate-erd --backend BD_Datos.accdb --access BD.accdb` → genera dos ficheros ERD
- `generate-erd` con tablas vinculadas ODBC → documenta DSN/Driver/Server sin comprobar alcanzabilidad
- `generate-erd` con tablas vinculadas a fichero no existente → sección "Orígenes no alcanzados"
- `end` → export final + resumen con módulos tocados
- `list` → muestra tabla con nombre, tipo y líneas de todos los módulos
- `delete Utilidades` → el módulo desaparece de la BD, archivos de src/ intactos
- `delete Utilidades --delete-src` → el módulo desaparece de la BD y de src/
- `delete Form_FormViejo --delete-src` → formulario borrado de BD y archivos `.form.txt` + `.cls` eliminados
- `rename modViejo modNuevo` → el módulo cambia de nombre en la BD y los archivos de src/ se renombran
- `rename Form_FormViejo Form_FormNuevo` → formulario renombrado en BD y archivos renombrados en src/
- `export-all --access "D:\otra\ruta\BD.accdb" --destination_root "D:\exports"` → funciona con rutas fuera de CWD
