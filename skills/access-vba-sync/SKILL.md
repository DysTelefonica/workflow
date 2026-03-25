# SKILL.md — access-vba-sync

## Objetivo

Sincronización bidireccional entre código VBA de Microsoft Access y archivos de texto en `src/`. Permite que una IA o un desarrollador edite archivos `.bas`, `.cls` y `.form.txt` y los importe a la BD, o exporte desde la BD para editarlos.

---

## 🚨 REGLAS ABSOLUTAS (leer ANTES de ejecutar nada)

### 1. SIEMPRE usar el CLI — NUNCA llamar al PS1 directamente

```powershell
# ✅ CORRECTO
node cli.js import MiModulo

# ❌ PROHIBIDO
.\VBAManager.ps1 -Action Import ...
powershell.exe -File VBAManager.ps1 ...
```

### 2. NUNCA ejecutar PowerShell suelto para verificar Access

El CLI gestiona Access internamente. Ejecutar `Get-Process` con `$_` desde shells no-PowerShell (Git Bash, WSL, terminal de Trae/OpenCode) causa bucles infinitos por interpolación de `$_` como variable bash.

```powershell
# ❌ PROHIBIDO — causa error /usr/bin/bash.ProcessName en bucle
Get-Process | Where-Object { $_.ProcessName -like '*ACCESS*' }

# ✅ Si realmente necesitas matar Access (tras un fallo), UNA sola vez:
powershell.exe -NoProfile -NonInteractive -Command 'Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force'
```

### 3. UN comando a la vez, ESPERAR a que termine

Cada `node cli.js <comando>` abre Access COM headless, opera, y cierra. Tiene timeout de 5 minutos.

```powershell
# ✅ Secuencial — esperar a que termine cada uno
node cli.js import ModuloA
node cli.js import ModuloB

# ❌ NUNCA en paralelo
node cli.js import ModuloA & node cli.js import ModuloB
```

Se pueden importar **varios módulos en un solo comando** (preferible):

```powershell
node cli.js import ModuloA ModuloB ModuloC
```

### 4. Encoding: archivos en `src/` son UTF-8 sin BOM

Al escribir o editar archivos en `src/`, guardar **siempre como UTF-8 sin BOM**. El CLI convierte automáticamente a ANSI (Windows-1252) al importar a Access, y de ANSI a UTF-8 al exportar. Los caracteres españoles (tildes, eñes, etc.) se preservan correctamente si se respeta esta regla.

Si ves caracteres corruptos (`S?` en vez de `Sí`):

```powershell
node cli.js fix-encoding --location Both
```

### 5. La BD debe estar CERRADA antes de cualquier comando

Si Access tiene la BD abierta, el CLI fallará. Cerrar Access antes de operar.

---

## Importación diferenciada por tipo de cambio

| Qué cambié | Comando | Qué importa |
|------------|---------|-------------|
| Código VBA de formulario (sin tocar UI) | `node cli.js import Form_X` | Solo `.cls` |
| UI de formulario (controles, layout) | `node cli.js import-form Form_X` | `.form.txt` (UI + código) |
| Módulo `.bas` o clase `.cls` | `node cli.js import MiModulo` | El archivo `.bas`/`.cls` |

**Regla clave:** si solo cambiaste código VBA de un formulario, usar `import` (no `import-form`). Usar `import-form` reimporta toda la UI y puede pisar cambios de layout hechos en Access.

---

## Estructura de `src/`

```
src/
├── modules/                    # Módulos estándar (.bas)
│   └── MiModulo.bas
├── classes/                    # Clases VBA (.cls)
│   └── CUsuario.cls
└── forms/                     # Formularios (DOS archivos por formulario)
    ├── Form_MiForm.form.txt   # UI completa + código (SaveAsText)
    └── Form_MiForm.cls        # Solo código VBA (para diff y edición)
```

---

## Comandos

### Exportar

```powershell
node cli.js export MiModulo                    # Un módulo (.bas/.cls)
node cli.js export ModA ModB ModC              # Varios módulos
node cli.js export-form Form_MiForm            # Un formulario (.form.txt + .cls)
node cli.js export-form Form_A Form_B Form_C   # Varios formularios
node cli.js export-all                         # Todos los módulos
node cli.js start                              # Export-all + inicia sesión
```

### Importar código (`.bas`/`.cls` — NO toca `.form.txt`)

```powershell
node cli.js import MiModulo                    # Un módulo
node cli.js import ModA ModB ModC              # Varios módulos
node cli.js import Form_MiForm                 # Código de formulario (solo .cls)
node cli.js import Form_A Form_B               # Código de varios formularios
node cli.js import-all                         # Todos los .bas/.cls de src/
```

### Importar formularios (`.form.txt` — UI + código)

```powershell
node cli.js import-form Form_MiForm            # Un formulario completo
node cli.js import-form Form_A Form_B Form_C   # Varios formularios
node cli.js import-form-all                    # Todos los .form.txt de src/
```

### Eliminar módulos de la BD

```powershell
node cli.js delete-module MiModulo             # Elimina de la BD (NO de src/)
node cli.js delete-module ModA ModB            # Varios a la vez
```

### Utilidades

```powershell
node cli.js fix-encoding                       # Corrige encoding en src/ y BD
node cli.js fix-encoding --location Src        # Solo archivos en src/
node cli.js fix-encoding --location Access     # Solo en la BD
node cli.js fix-encoding ModA ModB             # Solo módulos específicos
node cli.js generate-erd                       # Genera docs/ERD/NombreBackend.md
node cli.js generate-erd --backend "X.accdb"   # Backend específico
node cli.js status                             # Estado de sesión
node cli.js end                                # Cierra sesión (export final)
node cli.js end --auto_export_on_end false      # Sin export final
node cli.js watch                              # Auto-sync al guardar archivos
```

---

## Workflow típico con IA

### Cambiar código VBA de un formulario

```powershell
# 1. Exportar para tener versión actualizada
node cli.js export-form Form_MiForm

# 2. Editar src/forms/Form_MiForm.cls (solo código)

# 3. Importar SOLO código — NO la UI
node cli.js import Form_MiForm
```

### Cambiar la UI de un formulario

```powershell
# 1. Exportar para tener .form.txt limpio
node cli.js export-form Form_MiForm

# 2. Editar src/forms/Form_MiForm.form.txt (UI + código)

# 3. Importar formulario completo
node cli.js import-form Form_MiForm
```

### Cambiar módulos/clases

```powershell
# 1. Exportar
node cli.js export MiModulo

# 2. Editar src/modules/MiModulo.bas o src/classes/MiClase.cls

# 3. Importar
node cli.js import MiModulo
```

### Tras cada importación

Abrir Access → VBE → Debug → Compile. Obligatorio para validar.

---

## Flags

| Flag | Descripción | Default |
|------|-------------|---------|
| `--access <ruta>` | Ruta a la BD (.accdb/.mdb) | Autodetecta en CWD |
| `--password <pwd>` | Contraseña de la BD | — |
| `--destination_root <dir>` | Carpeta de trabajo | `src` |
| `--debounce_ms <n>` | Debounce en ms para watch | `600` |
| `--location Both\|Src\|Access` | Ámbito de fix-encoding | `Both` |
| `--backend <ruta>` | Backend para generate-erd | Autodetecta `*_Datos.accdb` |
| `--erd_path <dir>` | Carpeta de salida ERD | `docs/ERD` |
| `--auto_export_on_end false` | Desactiva export al cerrar | `true` |

---

## Protocolo de recuperación ante fallos

### El comando falló con error

```powershell
# 1. Matar Access si quedó abierto
powershell.exe -NoProfile -NonInteractive -Command 'Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force'

# 2. Esperar a que se libere el lock
Start-Sleep -Seconds 2

# 3. Reintentar
node cli.js import MiModulo
```

### El comando se cuelga (más de 2 minutos sin output)

El timeout automático (5 min) lo matará. Si necesitas intervenir antes:

```powershell
# 1. Ctrl+C en la terminal (o matar el proceso node)

# 2. Matar Access huérfano
powershell.exe -NoProfile -NonInteractive -Command 'Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force'

# 3. Esperar y reintentar
Start-Sleep -Seconds 2
node cli.js import MiModulo
```

### Error "UnauthorizedAccess" / "AuthorizationManager"

Si aparece un error de `ExecutionPolicy`, es un problema de GPO. El CLI ya usa `-ExecutionPolicy Bypass` con `-Command` para evitarlo. Si persiste, ejecutar manualmente:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Y luego hacer `Unblock-File` sobre el `.ps1`:

```powershell
Unblock-File -Path ".agents\skills\access-vba-sync\VBAManager.ps1"
```

### Caracteres corruptos (tildes, eñes)

```powershell
node cli.js fix-encoding --location Both
```

### "Módulo no encontrado" / "Subíndice fuera del intervalo"

Para formularios, siempre usar el nombre con prefijo `Form_`: `Form_MiFormulario`, no `MiFormulario`.

---

## Resumen rápido

| Necesito | Comando |
|----------|---------|
| Exportar módulo(s) | `node cli.js export ModA ModB` |
| Exportar formulario(s) | `node cli.js export-form Form_A Form_B` |
| Importar código (formulario) | `node cli.js import Form_A Form_B` |
| Importar UI formulario(s) | `node cli.js import-form Form_A Form_B` |
| Importar módulo/clase(s) | `node cli.js import ModA ModB` |
| Eliminar módulo(s) de la BD | `node cli.js delete-module ModA ModB` |
| Corregir encoding | `node cli.js fix-encoding` |
| Generar ERD | `node cli.js generate-erd` |
| Auto-sync | `node cli.js watch` |
| Estado | `node cli.js status` |
| Cerrar sesión | `node cli.js end` |
