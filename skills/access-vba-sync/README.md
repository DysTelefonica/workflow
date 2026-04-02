# VBA Access Sync

Skill de sincronización bidireccional para proyectos Microsoft Access/VBA.

## Descripción

Permite trabajar con el código VBA de Microsoft Access como archivos de texto plano en tu editor favorito, manteniendo sincronización automática entre el código fuente y la base de datos.

### Características Principales

- **Export desatendido**: Extrae todos los módulos VBA sin ejecutar formularios de inicio ni macros AutoExec
- **Export/Import selectivo**: Opera sobre módulos individuales por nombre
- **Estructura organizada**: Separa automáticamente en carpetas según el tipo (`modules/`, `classes/`, `forms/`)
- **Sincronización automática**: Detecta cambios en archivos y los importa a Access en tiempo real
- **Borrado de módulos**: Elimina módulos de la BD y opcionalmente de `src/`
- **Renombrado de módulos**: Renombra en la BD y actualiza los archivos de `src/` automáticamente
- **Listado de módulos**: Muestra todos los módulos VBA con su tipo y líneas de código
- **Fix-encoding**: Corrige archivos con BOM UTF-8 o encoding incorrecto, en src/ o en la BD
- **Generación de ERD**: Exporta la estructura de tablas a Markdown — detecta tablas vinculadas (Access, ODBC, Excel, SharePoint, Text/CSV) con clasificación por tipo y comprobación de alcanzabilidad
- **Sesiones persistentes**: Mantiene estado entre ejecuciones
- **BD fuera de CWD**: Acepta rutas absolutas con `--access` sin restricción de directorio
- **Destino configurable**: `--destination_root` permite elegir dónde van los archivos exportados

## Requisitos

- Windows 10/11
- Microsoft Access (versión 2016 o superior recomendada)
- PowerShell 5.1+
- Node.js 18+

## Instalación

```powershell
cd access-vba-sync
npm install
```

## Estructura de Archivos Exportados

```
src/
├── modules/          # Módulos estándar (.bas)
│   └── MiModulo.bas
├── classes/          # Clases VBA (.cls)
│   └── CUsuario.cls
└── forms/            # Formularios Access
    ├── Form_MiForm.form.txt   # UI + código completo (vía SaveAsText)
    └── Form_MiForm.cls        # Solo código VBA (para diff y lectura rápida)
```

La carpeta de destino es configurable:

```powershell
# Usar una carpeta distinta de src/
node cli.js start --destination_root "D:\exports\mi_proyecto"
```

## Comandos

### `start` — Iniciar sesión

```powershell
node cli.js start [--access <ruta>] [--destination_root <dir>] [--password <pwd>]
```

Exporta todos los módulos de la BD hacia `src/` e inicia la sesión. Punto de partida de cualquier sesión de trabajo.

---

### `watch` — Sincronización automática

```powershell
node cli.js watch [--access <ruta>] [--destination_root <dir>] [--debounce_ms <n>] [--password <pwd>]
```

Inicia sesión (si no hay una activa) y monitoriza cambios en `src/`. Al guardar un archivo `.bas`, `.cls`, `.frm` o `.form.txt`, lo importa automáticamente a la BD con debounce.

---

### `export <Mod...>` — Exportar módulos específicos

```powershell
node cli.js export Form_FormInicial Utilidades [--access <ruta>] [--password <pwd>]
```

Exporta uno o más módulos por nombre desde la BD hacia `src/`. El nombre es el del VBComponent (con prefijo `Form_` para formularios). No requiere sesión activa.

---

### `export-all` — Exportar todos los módulos

```powershell
node cli.js export-all [--access <ruta>] [--destination_root <dir>] [--password <pwd>]
```

Exporta todos los módulos de la BD hacia `src/`. Equivalente a `start` pero sin iniciar sesión.

---

### `import <Mod...>` — Importar módulos específicos

```powershell
node cli.js import Form_FormInicial Utilidades [--access <ruta>] [--password <pwd>]
```

Importa uno o más módulos por nombre desde `src/` hacia la BD. Tras importar recuerda compilar en el VBE.

---

### `import-all` — Importar todos los módulos

```powershell
node cli.js import-all [--access <ruta>] [--destination_root <dir>] [--password <pwd>]
```

Importa todos los archivos de `src/` hacia la BD.

---

### `sync <Mod...>` — Alias de import

```powershell
node cli.js sync Utilidades Validaciones
```

Alias de `import`, mantenido por compatibilidad.

---

### `delete <Mod...>` — Borrar módulos

```powershell
node cli.js delete Utilidades Form_FormViejo [--access <ruta>] [--password <pwd>] [--delete-src]
```

Borra uno o más módulos de la BD. Cada módulo se procesa individualmente — los errores en uno no detienen el lote.

**`--delete-src`**: si se pasa, también borra los archivos correspondientes de `src/` (`.form.txt`, `.cls`, `.bas`). Sin este flag, los archivos de `src/` permanecen intactos.

```powershell
# Solo borrar de la BD
node cli.js delete modObsoleto

# Borrar de la BD y de src/
node cli.js delete modObsoleto --delete-src
```

---

### `rename <Old> <New>` — Renombrar un módulo

```powershell
node cli.js rename modNombreViejo modNombreNuevo [--access <ruta>] [--password <pwd>]
```

Renombra un módulo en la BD y actualiza automáticamente los archivos en `src/`. Solo se puede renombrar un módulo a la vez.

```powershell
# Renombrar un módulo estándar
node cli.js rename Utilidades UtilidadesV2

# Renombrar un formulario
node cli.js rename Form_FormViejo Form_FormNuevo
```

**Nota**: el rename no actualiza referencias al módulo en otros módulos. Hay que buscar y reemplazar manualmente.

---

### `list` — Listar módulos

```powershell
node cli.js list [--access <ruta>] [--password <pwd>]
```

Muestra una tabla con todos los módulos VBA de la BD: nombre, tipo (Module/Class/Form) y número de líneas de código. Los tipos se muestran con colores distintos. No requiere sesión activa.

---

### `import-form <Mod...>` — Importar formularios (UI + código)

```powershell
node cli.js import-form Form_frmDatosPC [--access <ruta>] [--password <pwd>]
```

Importa formularios forzando el uso de `.form.txt`/`.frm` (layout completo con UI + código).

---

### `import-code <Mod...>` — Importar solo code-behind

```powershell
node cli.js import-code Form_frmDatosPC [--access <ruta>] [--password <pwd>]
```

Importa solo el código VBA desde `.cls`/`.bas`, sin tocar el layout del formulario.

---

### `fix-encoding` — Corregir encoding

```powershell
node cli.js fix-encoding [<Mod...>] [--access <ruta>] [--location Both|Src|Access] [--password <pwd>]
```

Corrige archivos con BOM UTF-8 u otros problemas de encoding. Sin módulos procesa todos; con módulos solo los indicados.

**`--location`** controla dónde se aplica:

| Valor | Efecto |
|---|---|
| `Both` (default) | Corrige en `src/` y en la BD |
| `Src` | Solo corrige los archivos en `src/` |
| `Access` | Solo reimporta desde `src/` para corregir en la BD |

---

### `generate-erd` — Generar documentación de tablas

```powershell
node cli.js generate-erd [--access <ruta>] [--backend <ruta>] [--erd_path <dir>] [--password <pwd>]
```

Extrae la estructura de tablas, campos, tipos, claves primarias, relaciones y tablas vinculadas a un archivo Markdown.

**Modos de uso:**

```powershell
# Solo backend (comportamiento clásico) — auto-detecta *_Datos.accdb
node cli.js generate-erd

# Solo backend explícito
node cli.js generate-erd --backend "BD_Datos.accdb"

# Solo frontend — muestra tablas vinculadas clasificadas por tipo
node cli.js generate-erd --access "MiFrontend.accdb"

# Frontend + backend — genera dos ficheros .md
node cli.js generate-erd --access "MiFrontend.accdb" --backend "BD_Datos.accdb"
```

Si no se pasa `--backend` ni `--access`, auto-detecta `*_Datos.accdb` en CWD. Si no lo encuentra, usa el frontend (sesión activa o auto-detect).

**Detección de tablas vinculadas:**

Al generar el ERD de una BD que contiene tablas vinculadas, el skill las detecta usando `TableDef.Attributes` y `TableDef.Connect`, y las clasifica por tipo de conexión:

| Tipo | Patrón `Connect` | Ejemplo |
|---|---|---|
| Access | `;DATABASE=ruta.accdb` | Backend local vinculado |
| ODBC | `ODBC;DSN=...` o `DRIVER=...` | SQL Server, Oracle, MySQL |
| Excel | `Excel 8.0;DATABASE=...` | Hoja de cálculo vinculada |
| SharePoint | `WSS;` o `SharePoint` | Lista de SharePoint |
| Text/CSV | `Text;DATABASE=...` | Fichero de texto/CSV |
| HTML | `HTML;DATABASE=...` | Tabla HTML |
| Otro | Cualquier otro formato | Catch-all |

Para cada tabla vinculada se extrae también `SourceTableName` (el nombre de la tabla en el origen cuando difiere del nombre local).

**Formato de salida del ERD:**

El Markdown generado incluye:

- Encabezado con desglose: `## Tablas (45 total: 12 locales, 33 vinculadas)`
- Cada tabla vinculada marcada: `### tblClientes _(vinculada: ODBC, origen: dbo.Clientes)_`
- Sección `## Relaciones` con las relaciones DAO
- Sección `## Tablas vinculadas` agrupada por tipo de conexión con tabla local, tabla origen y destino
- Subsección `### Orígenes no alcanzados` para ficheros que no existen en la ruta
- Subsección `### Conexiones ODBC` con DSN/Driver/Server documentados

El fichero se nombra igual que la BD: `NombreBaseDatos.md` dentro de `--erd_path` (default: `docs/ERD/`).

---

### `status` — Estado de la sesión

```powershell
node cli.js status
```

Muestra el estado de la sesión activa: BD, rutas, módulos tocados, pendientes y último sync.

---

### `end` — Cerrar sesión

```powershell
node cli.js end [--auto_export_on_end false]
```

Para el watcher si está activo, importa pendientes, realiza el export final (configurable) y restaura la configuración de Access (StartupForm, AutoExec, AllowBypassKey).

---

## Flags globales

| Flag | Descripción | Default |
|---|---|---|
| `--access <ruta>` | Ruta a la BD (.accdb/.accde/.mdb/.mde). Acepta rutas absolutas fuera de CWD. También sirve para generate-erd (frontend). | Autodetecta en CWD |
| `--password <pwd>` | Contraseña de la BD si está protegida | — |
| `--destination_root <dir>` | Carpeta de trabajo (export/import). Acepta rutas absolutas. | `src` |
| `--debounce_ms <n>` | Espera en ms antes de importar en watch | `600` |
| `--location Both\|Src\|Access` | Ámbito de fix-encoding | `Both` |
| `--backend <ruta>` | Backend para generate-erd. Se puede combinar con `--access` para generar ambos. | Autodetecta `*_Datos.accdb` |
| `--erd_path <dir>` | Carpeta de salida para el ERD | `docs/ERD` |
| `--auto_export_on_end false` | Desactiva export final al cerrar | `true` |
| `--delete-src` | Para delete: también borrar archivos de src/ | `false` |

---

## Comportamiento interno

El skill gestiona automáticamente antes de abrir la BD:

- **StartupForm**: lo deshabilita temporalmente para evitar que arranque el formulario de inicio
- **AutoExec**: renombra la macro para evitar su ejecución
- **AllowBypassKey**: asegura acceso completo

Al cerrar la sesión restaura todos estos valores al estado original.

Access se abre en modo headless (`Visible=false`, `UserControl=false`) — no aparece ninguna ventana ni diálogo. **La BD debe estar cerrada antes de ejecutar cualquier comando.**

---

## Integración con el framework VBA-SDD

Para desplegar el skill en un nuevo proyecto:

```powershell
.\deploy.ps1
```

Esto copia el skill, genera la estructura inicial y exporta el código VBA automáticamente.

---

## Flujo de trabajo típico

```powershell
# 1. Ver qué módulos tiene la BD
node cli.js list

# 2. Inicio de sesión — snapshot inicial
node cli.js start

# 3. (Opcional) Generar contexto de datos para la IA
node cli.js generate-erd                              # backend (auto-detecta *_Datos.accdb)
node cli.js generate-erd --access MiFrontend.accdb    # frontend (tablas vinculadas)

# 4a. Modo automático: la IA edita src/ y los cambios se importan solos
node cli.js watch

# 4b. Modo manual: importar después de que la IA termine
node cli.js import Form_FormGestion ModuloDAO

# 5. Compilar en Access
# Abre Access → VBE → Debug → Compile

# 6. Limpiar módulos obsoletos
node cli.js delete modViejo --delete-src

# 7. Renombrar si hace falta
node cli.js rename modTemporal modDefinitivo

# 8. Cerrar sesión
node cli.js end
```

---

## Resolución de Problemas

### Access se abre visiblemente
Hay una instancia previa de Access en ejecución. Ciérrala antes:
```powershell
Get-Process MSACCESS | Stop-Process -Force
```

### Error "archivo en uso" o "locked"
La BD está abierta en otro proceso. Cierra Access completamente antes de ejecutar el skill.

### Los formularios se exportan sin la UI (solo código)
El export de formularios usa `SaveAsText` que requiere que el objeto Access sea accesible. Si falla, revisa que la BD no esté bloqueada y que `--access` apunte al frontend correcto.

### Encoding incorrecto en tildes o caracteres especiales
```powershell
node cli.js fix-encoding --location Both
```

### La BD tiene contraseña
Pasa siempre `--password "micontraseña"` en todos los comandos.

### Error al borrar un formulario
Los formularios se borran con `DoCmd.DeleteObject`, no con `VBComponents.Remove`. Si el error persiste, asegúrate de que el formulario no está abierto en otra instancia de Access.

### Rename no actualiza referencias
El rename solo cambia el nombre del módulo en la BD y los archivos en src/. Las llamadas a ese módulo desde otros módulos deben actualizarse manualmente (buscar y reemplazar).

### La BD está fuera de CWD
Usa `--access` con la ruta completa:
```powershell
node cli.js list --access "D:\proyectos\MiBD.accdb"
node cli.js export-all --access "D:\proyectos\MiBD.accdb" --destination_root "D:\exports"
```

### El ERD no muestra tablas vinculadas
Las tablas vinculadas solo aparecen si generas el ERD del **frontend** (que es donde viven los vínculos). Usa `--access`:
```powershell
node cli.js generate-erd --access "MiFrontend.accdb"
```
El backend (`*_Datos.accdb`) tiene tablas locales — no vinculadas.

### El ERD muestra "Orígenes no alcanzados"
El fichero al que apunta la vinculación no existe en la ruta esperada. Esto es normal si generas el ERD desde una máquina distinta a donde se usa la BD, o si el backend está en una unidad de red no conectada.
