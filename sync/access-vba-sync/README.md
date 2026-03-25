# VBA Access Sync

Skill de sincronización bidireccional para proyectos Microsoft Access/VBA.

## Descripción

Permite trabajar con el código VBA de Microsoft Access como archivos de texto plano en tu editor favorito, manteniendo sincronización automática entre el código fuente y la base de datos.

### Características Principales

- **Export desatendido**: Extrae todos los módulos VBA sin ejecutar formularios de inicio ni macros AutoExec
- **Export/Import selectivo**: Opera sobre módulos individuales por nombre
- **Estructura organizada**: Separa automáticamente en carpetas según el tipo (`modules/`, `classes/`, `forms/`)
- **Sincronización automática**: Detecta cambios en archivos y los importa a Access en tiempo real
- **Fix-encoding**: Corrige archivos con BOM UTF-8 o encoding incorrecto, en src/ o en la BD
- **Generación de ERD**: Exporta la estructura de la base de datos backend a Markdown
- **Sesiones persistentes**: Mantiene estado entre ejecuciones

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

Exporta uno o más módulos por nombre desde la BD hacia `src/`. El nombre es el del VBComponent (con prefijo `Form_` para formularios).

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
node cli.js generate-erd [--backend <ruta>] [--erd_path <dir>] [--password <pwd>]
```

Extrae la estructura de tablas, campos, tipos, claves primarias y relaciones del backend a un archivo Markdown. Autodetecta `*_Datos.accdb` en el directorio actual si no se especifica `--backend`.

El archivo generado se nombra igual que el backend: `NombreBackend.md` dentro de `--erd_path` (default: `docs/ERD/`).

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
| `--access <ruta>` | Ruta a la BD (.accdb/.accde/.mdb/.mde) | Autodetecta en CWD |
| `--password <pwd>` | Contraseña de la BD si está protegida | — |
| `--destination_root <dir>` | Carpeta de trabajo (export/import) | `src` |
| `--debounce_ms <n>` | Espera en ms antes de importar en watch | `600` |
| `--location Both\|Src\|Access` | Ámbito de fix-encoding | `Both` |
| `--backend <ruta>` | Backend para generate-erd | Autodetecta `*_Datos.accdb` |
| `--erd_path <dir>` | Carpeta de salida para el ERD | `docs/ERD` |
| `--auto_export_on_end false` | Desactiva export final al cerrar | `true` |

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
# 1. Inicio de sesión — snapshot inicial
node cli.js start

# 2. (Opcional) Generar contexto de datos para la IA
node cli.js generate-erd

# 3a. Modo automático: la IA edita src/ y los cambios se importan solos
node cli.js watch

# 3b. Modo manual: importar después de que la IA termine
node cli.js import Form_FormGestion ModuloDAO

# 4. Compilar en Access
# Abre Access → VBE → Debug → Compile

# 5. Cerrar sesión
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
