# Plan de Trabajo: Corrección de Bugs Críticos

## TL;DR

> **Objetivo**: Corregir los bugs que afectan a usuarios reales del framework Dysflow.
> 
> **Deliverables**:
> - `dysflow plan <num>` funciona correctamente
> - Estructura de carpetas coincide con documentación
> - Manejo de errores mejorado en CLI
> 
> **Esfuerzo**: Bajo (3 archivos, cambios focalizados)
> **Ejecución**: Secuencial (dependencias mínimas)

---

## Context

### Problema Identificado
El análisis del repositorio reveló 2 bugs que **AFECTAN A USUARIOS REALES**:

1. **B1 (CRÍTICO)**: `dysflow plan <num>` busca `docs/plans/active` que no existe
   - Error: `ENOENT: no such file or directory`
   - Comando documentado en onboarding.html línea 1038

2. **D2**: Inconsistencia de rutas entre documentación e implementación
   - Docs dicen: `src/clases/`, `src/modulos/`, `src/formularios/`
   - init-access.js crea: `src/classes/`, `src/modules/`, `src/forms/`

---

## Work Objectives

### Objetivo Principal
Que todos los comandos CLI documentados en onboarding.html funcionen correctamente.

### Entregables
- [ ] Directorio `docs/plans/active/` y `docs/plans/completed/` se crea en `dysflow init`
- [ ] `dysflow plan <num>` encuentra specs y crea ramas
- [ ] Estructura de carpetas usa español (`clases`, `modulos`, `formularios`)
- [ ] Errores CLI devuelven códigos de salida correctos

---

## Execution Strategy

### Análisis de Cambios

| Archivo | Cambio Requerido |
|---------|------------------|
| `installers/init-access.js` | Agregar `docs/plans/active` y `docs/plans/completed` + corregir rutas |
| `cli/commands/plan-create.js` | Mejorar mensaje de error si no encuentra plan |
| `cli/commands/spec-create.js` | Cambiar `process.exit()` → `process.exit(1)` |

### Dependencias
- Ninguna dependencia externa
- Cambios en archivos independientes

---

## TODOs

### Tarea 1: Crear directorios de planes en init-access.js

**Qué hacer**:
1. Leer `installers/init-access.js`
2. Agregar `docs/plans/active` y `docs/plans/completed` al array de directorios (líneas 187-200)
3. Corregir rutas de `src/` para usar español:
   - `src/modules` → `src/modulos`
   - `src/classes` → `src/clases`
   - `src/forms` → `src/formularios`

**QA Scenarios**:

```
Scenario: dysflow init crea estructura correcta
  Tool: Bash
  Preconditions: Ninguna
  Steps:
    1. cd a carpeta temporal
    2. Ejecutar: node ../workflow/cli/workflow.js init access
    3. Verificar que existen: docs/plans/active, docs/plans/completed
    4. Verificar que existen: src/clases, src/modulos, src/formularios
  Expected Result: Todos los directorios existen
  Evidence: .sisyphus/evidence/init-estructura.txt

Scenario: dysflow init NO crea carpetas en inglés
  Tool: Bash
  Preconditions: Ninguna
  Steps:
    1. Verificar que NO existen: src/classes, src/modules, src/forms
  Expected Result: Archivos no existen (o si existen, son los correctos en español)
  Evidence: .sisyphus/evidence/init-no-english.txt
```

**Commit**: YES
- Message: `fix(init): agregar docs/plans y corregir rutas a español`
- Files: `installers/init-access.js`

---

### Tarea 2: Mejorar plan-create.js

**Qué hacer**:
1. Leer `cli/commands/plan-create.js`
2. Agregar validación si el directorio no existe:
   ```javascript
   const plansDir = "docs/plans/active"
   if (!fs.existsSync(plansDir)) {
     console.log("Error: El directorio docs/plans/active no existe.")
     console.log("Ejecuta 'dysflow init access' primero.")
     process.exit(1)
   }
   const folders = fs.readdirSync(plansDir)
   ```
3. Agregar mensaje de error claro si no encuentra el plan

**QA Scenarios**:

```
Scenario: dysflow plan con directorio inexistente
  Tool: Bash
  Preconditions: Proyecto sin docs/plans/active
  Steps:
    1. Ejecutar: node cli/workflow.js plan 1
  Expected Result: Mensaje claro + exit code 1
  Evidence: .sisyphus/evidence/plan-no-dir.txt

Scenario: dysflow plan con plan inexistente
  Tool: Bash
  Preconditions: Proyecto con docs/plans/active vacío
  Steps:
    1. Ejecutar: node cli/workflow.js plan 999
  Expected Result: "Plan not found" + exit code 1
  Evidence: .sisyphus/evidence/plan-no-existe.txt
```

**Commit**: YES
- Message: `fix(plan-create): validar existencia de directorio y plan`
- Files: `cli/commands/plan-create.js`

---

### Tarea 3: Corregir spec-create.js exit code

**Qué hacer**:
1. Leer `cli/commands/spec-create.js`
2. Cambiar línea 11:
   ```javascript
   // Antes
   process.exit()
   
   // Después
   process.exit(1)
   ```

**QA Scenarios**:

```
Scenario: dysflow spec con spec inexistente devuelve código de error
  Tool: Bash
  Preconditions: Ninguna
  Steps:
    1. Ejecutar: node cli/workflow.js spec 9999; echo "Exit code: $?"
  Expected Result: Exit code 1 (no 0)
  Evidence: .sisyphus/evidence/spec-exit-code.txt
```

**Commit**: YES
- Message: `fix(spec-create): devolver código de error cuando no encuentra spec`
- Files: `cli/commands/spec-create.js`

---

### Tarea 4: Mejorar git.js con manejo de errores

**Qué hacer**:
1. Leer `cli/utils/git.js`
2. Agregar try-catch:
   ```javascript
   const { execSync } = require("child_process")
   
   function git(cmd) {
     try {
       return execSync(`git ${cmd}`, { encoding: "utf8", stdio: "pipe" }).trim()
     } catch (err) {
       if (err.status) {
         console.error(`Git error: ${err.message}`)
         process.exit(1)
       }
       throw new Error(`Git no disponible o no es un repositorio: ${err.message}`)
     }
   }
   
   module.exports = { git }
   ```

**QA Scenarios**:

```
Scenario: git command fuera de repositorio
  Tool: Bash
  Preconditions: Carpeta sin .git
  Steps:
    1. Crear carpeta temporal sin git
    2. Ejecutar: node cli/workflow.js spec 1
  Expected Result: Mensaje de error claro, no stack trace
  Evidence: .sisyphus/evidence/git-no-repo.txt
```

**Commit**: YES
- Message: `fix(git): agregar manejo de errores y códigos de salida`
- Files: `cli/utils/git.js`

---

## Final Verification Wave

- [ ] **Estructura creada**: `docs/plans/active/` y `docs/plans/completed/` existen tras init
- [ ] **Rutas corregidas**: `src/clases/`, `src/modulos/`, `src/formularios/` existen
- [ ] **dysflow plan funciona**: Crea rama git correctamente
- [ ] **Códigos de error**: spec/plan fallidos devuelven exit code 1
- [ ] **Errores claros**: Mensajes útiles, no stack traces

---

## Commit Strategy

| # | Message | Files |
|---|---------|-------|
| 1 | `fix(init): agregar docs/plans y corregir rutas a español` | `installers/init-access.js` |
| 2 | `fix(plan-create): validar existencia de directorio y plan` | `cli/commands/plan-create.js` |
| 3 | `fix(spec-create): devolver código de error cuando no encuentra spec` | `cli/commands/spec-create.js` |
| 4 | `fix(git): agregar manejo de errores y códigos de salida` | `cli/utils/git.js` |

---

## Success Criteria

```bash
# Verificar que todo funciona
node cli/workflow.js init access  # Crea estructura correcta
node cli/workflow.js plan 1       # Funciona (o dice que no hay planes)
node cli/workflow.js spec 999    # Exit code 1
```

- [ ] Todos los tests pasan
- [ ] Estructura coincide con documentación (onboarding.html)
- [ ] Códigos de salida correctos en errores
