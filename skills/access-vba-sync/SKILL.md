# SKILL.md — Skill para workflow Access/VBA

## Objetivo

Automatizar el workflow de desarrollo entre código VBA de Microsoft Access y archivos de texto en disco, permitiendo que una IA (o un desarrollador) edite el código fuentes y lo sincronice con la BD.

---

## ⚠️ REGLA CRÍTICA: IMPORTACIÓN DIFERENCIADA POR TIPO DE CAMBIO

| Tipo de cambio | Archivos a importar | Comando |
|----------------|---------------------|---------|
| **Código VBA de formulario** (sin UI) | Solo `.cls` | `node cli.js import Form_Nombre [Form_Otro ...]` |
| **UI de formulario** (controles, layout, propiedades) | `.cls` + `.form.txt` | `node cli.js import-form Form_Nombre [Form_Otro ...]` |
| **Módulo .bas o clase .cls** | Solo el archivo | `node cli.js import NombreModulo [Modulo2 ...]` |

> **Se pueden pasar VARIOS formularios o módulos a la vez** — el CLI los procesa todos en una sola ejecución.

**¿Por qué esta distinción?**
- `import` → solo código VBA (`.cls` de formulario o `.bas`/`.cls` normal)
- `import-form` → toda la UI del formulario (`.form.txt`) + código

**Si importás `.form.txt` cuando no tocaste la UI, podés perder cambios de layout hechos en Access.**

---

## ⚠️ REGLA DE ORO: SIEMPRE USAR EL CLI

**NUNCA llamar directamente al `VBAManager.ps1`.** 

Usar SIEMPRE el `cli.js` como interfaz:

```powershell
node cli.js <comando>
```

El CLI maneja correctamente el directorio de trabajo (`cwd`), la resolución de rutas y los parámetros.

**❌ MAL (no usar):**
```powershell
.\VBAManager.ps1 -Action Export ...
```

**✅ BIEN (siempre así):**
```powershell
node cli.js export MiModulo
```

---

## Estructura de archivos exportados

```
src/
├── modules/                    # Módulos estándar (.bas)
│   └── MiModulo.bas
├── classes/                    # Clases VBA (.cls)
│   └── CUsuario.cls
└── forms/                     # Formularios Access (DOS archivos por formulario)
    ├── Form_MiForm.form.txt   # UI + código completo (SaveAsText)
    └── Form_MiForm.cls        # Solo código VBA (para diff y edición)
```

**Importante:** Cada formulario genera DOS archivos:
- `.form.txt` → contiene UI (controles, propiedades) + código VBA
- `.cls` → contiene SOLO el código VBA

---

## Comandos esenciales (con ejemplos)

### Exportar módulos

```powershell
# Exportar UN módulo específico (código .bas o .cls)
node cli.js export MiModulo

# Exportar VARIOS módulos a la vez
node cli.js export ModuloA ModuloB ModuloC

# Exportar un formulario (genera .form.txt + .cls)
node cli.js export-form Form_MiFormulario

# Exportar VARIOS formularios a la vez
node cli.js export-form Form_FormGestion Form_FormDetalle Form_FormNC

# Exportar TODOS los módulos (start de sesión)
node cli.js start

# Exportar todo sin iniciar sesión
node cli.js export-all
```

### Importar módulos (código)

```powershell
# Importar código (.bas/.cls) de UN módulo — NO toca .form.txt
node cli.js import MiModulo

# Importar VARIOS módulos de código a la vez
node cli.js import ModuloA ModuloB ModuloC

# Importar código de UN formulario (.cls) — NO toca .form.txt
# Usa esto cuando solo cambiaste código VBA del formulario, NO la UI
node cli.js import Form_MiFormulario

# Importar código de VARIOS formularios a la vez
node cli.js import Form_MiFormulario Form_OtroFormulario

# Importar TODOS los módulos de código (no formularios)
node cli.js import-all
```

### Importar formularios (UI + código)

```powershell
# Importar formulario completo (.form.txt) — UI + código VBA
# Usa esto cuando cambiaste controles, propiedades, o addediste/removiste controles
node cli.js import-form Form_MiFormulario

# Importar VARIOS formularios completos a la vez
node cli.js import-form Form_FormGestion Form_FormDetalle Form_FormNC

# Importar TODOS los formularios
node cli.js import-form-all
```

---

## 🔑 CUANDO CAMBIAS CÓDIGO VBA DE UN FORMULARIO

Si solo editás el código VBA de un formulario (sin tocar UI):

```powershell
# 1. Exportás el formulario para tener la versión actualizada
node cli.js export-form Form_MiFormulario

# 2. Editás el .cls (código VBA)
#    src/forms/Form_MiFormulario.cls

# 3. Importás SOLO el código (NO la UI)
node cli.js import Form_MiFormulario
```

**❌ NO hacer:** `import-form` → esto reimporta la UI completa y puede perder cambios de layout hechos en Access.

---

## 🔑 CUANDO CAMBIAS LA UI DE UN FORMULARIO

Si agregaste/removiste controles, cambiaste propiedades de controles, o modificaste el layout:

```powershell
# 1. Exportás el formulario (para tener base .form.txt limpia)
node cli.js export-form Form_MiFormulario

# 2. Editás el .form.txt (UI + código)
#    src/forms/Form_MiFormulario.form.txt

# 3. Importás el formulario completo (UI + código)
node cli.js import-form Form_MiFormulario
```

---

## Workflow típico con IA

```
┌─────────────────────────────────────────────────────────────────┐
│ 1. INICIO: Exportar módulo a modificar                          │
│    node cli.js export-form Form_NCProyecto                      │
│    (o node cli.js export MiModulo si es un .bas)                │
└────────────────────────┬────────────────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│ 2. IA EDITORIAL: Modifica los archivos en src/                  │
│    - Código VBA → editar .cls (formularios) o .bas (módulos)   │
│    - UI formularios → editar .form.txt                          │
└────────────────────────┬────────────────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│ 3. SINCRONIZAR: Importar cambios a Access                       │
│                                                                 │
│    Si solo cambiaste CÓDIGO de formulario:                      │
│    → node cli.js import Form_NCProyecto                         │
│                                                                 │
│    Si cambiaste UI de formulario:                               │
│    → node cli.js import-form Form_NCProyecto                    │
│                                                                 │
│    Si cambiaste módulos/clases:                                 │
│    → node cli.js import MiModulo                                │
└────────────────────────┬────────────────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│ 4. COMPILAR: Abrir Access → VBE → Debug → Compile              │
│    (Obligatorio tras cada importación)                          │
└────────────────────────┬────────────────────────────────────────┘
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│ 5. PROBAR: Validar que los cambios funcionan en Access          │
└─────────────────────────────────────────────────────────────────┘
```

---

## Mode Watch (automático)

El modo watch detecta cambios en archivos y los importa automáticamente:

```powershell
node cli.js watch
```

A partir de entonces:
- Guardás un `.bas` → se importa automáticamente
- Guardás un `.cls` de formulario → se importa automáticamente como código
- Guardás un `.form.txt` → se importa automáticamente como formulario

---

## Sesión y estado

El skill mantiene estado en `.access-vba-skill/session.json`:

```powershell
node cli.js status    # Ver estado actual
node cli.js end       # Cerrar sesión (export final opcional)
```

---

## Generar documentación ERD

```powershell
# Genera docs/ERD/NombreBackend.md
node cli.js generate-erd

# Con backend específico
node cli.js generate-erd --backend "MiBackend.accdb" --erd_path "docs/ERD"
```

---

## Flags disponibles

| Flag | Descripción | Default |
|------|-------------|---------|
| `--access <ruta>` | Ruta a la BD (.accdb) | Autodetecta en CWD |
| `--password <pwd>` | Contraseña de la BD | — |
| `--destination_root <dir>` | Carpeta de trabajo | `src` |
| `--debounce_ms <n>` | Ms espera antes de importar en watch | `600` |

---

## Requisitos

- **Windows** con Microsoft Access instalado
- **La BD debe estar cerrada** antes de ejecutar cualquier comando
- El `cli.js` usa `cwd` del proceso Node → ejecutar desde la raíz del proyecto

---

## Resolución de problemas

### "Access está bloqueado"
La BD está abierta en Access. Cerrarla antes de usar el skill.

```powershell
Get-Process MSACCESS | Stop-Process -Force
```

### "Módulo no encontrado"
- Verificar que el nombre del módulo sea exacto (incluir prefijo `Form_` para formularios)
- Usar `node cli.js status` para ver módulos disponibles

### "Subíndice fuera del intervalo"
- Para formularios, usar el nombre con prefijo `Form_`: `Form_MiFormulario`, no `MiFormulario`

---

## Resumen rápido

| Qué necesito | Comando | ¿Qué importa? |
|--------------|---------|---------------|
| Exportar módulo(s) | `node cli.js export ModA ModB` | `.bas` o `.cls` |
| Exportar formulario(s) | `node cli.js export-form Form_A Form_B` | `.form.txt` + `.cls` |
| Importar código VBA (formulario) | `node cli.js import Form_A Form_B` | **Solo `.cls`** |
| Importar UI formulario(s) | `node cli.js import-form Form_A Form_B` | **`.form.txt` + `.cls`** |
| Importar módulo/clase(s) | `node cli.js import ModA ModB` | `.bas` o `.cls` |
| Sincronización automática | `node cli.js watch` | Auto-detecta tipo |
| Estado de sesión | `node cli.js status` | — |
| Cerrar sesión | `node cli.js end` | — |
