# Dysflow — Framework de Desarrollo Asistido por IA para Microsoft Access

> **Versión:** 1.0.0  
> **Stack:** Microsoft Access + VBA + Git + IA (Trae/OpenCode)  
> **Licencia:** MIT

---

## ¿Qué es Dysflow?

Dysflow es un **framework de desarrollo basado en especificaciones (SDD — Spec-Driven Development)** diseñado específicamente para proyectos **Microsoft Access con VBA**. Integra un flujo de trabajo estructurado con asistencia de inteligencia artificial para garantizar:

- **Trazabilidad total**: Cada cambio parte de una especificación documentada.
- **Cero regresiones**: Análisis de impacto antes de modificar código.
- **Entrega manual controlada**: El código se copia manualmente al editor VBA.
- **Calidad arquitectónica**: Separación MVVM, transacciones explícitas, gestión de errores.

### ¿Para quién es?

- Equipos que desarrollan aplicaciones Access/VBA en entornos corporativos.
- Desarrolladores que usan IA (Trae, OpenCode, Cursor) para acelerar el desarrollo.
- Técnicos de calidad que necesitan trazabilidad en sus proyectos.

---

## Instalación

### Requisitos previos

| Requisito | Versión mínima | Notas |
| :--- | :--- | :--- |
| **Windows** | 10/11 | Solo funciona en Windows |
| **Node.js** | 18+ | Para CLI de dysflow |
| **Git** | 2.30+ | Control de versiones |
| **Microsoft Access** | 2016+ | Entorno de ejecución y desarrollo |
| **PowerShell** | 7+ | Para scripts de automatización |
| **Trae / OpenCode** | — | Agente de IA (opcional) |

### Instalación global

```powershell
# Clonar el repositorio
git clone https://github.com/DysTelefonica/workflow.git dysflow-framework
cd dysflow-framework

# Instalar dependencias
npm install

# Instalar CLI globalmente (opcional)
npm link
```

### Verificar instalación

```powershell
dysflow --help
```

Deberías ver:

```
Usage: dysflow [options] [command]

Commands:
  spec <number>              Create spec branch
  plan <number>              Create plan branch
  release                    Create release from develop
  hotfix <name>              Create hotfix branch
  changelog <from>           Generate changelog
  next-release               Get next release tag
  init <type>                Initialize project (access)
  update                     Update skills and rules
```

---

## Inicializar un proyecto nuevo

### Paso 1: Crear estructura base

```powershell
dysflow init access
```

Esto crea la estructura de carpetas estándar:

```
MiProyecto/
├── .agent/
│   ├── AGENTS.md           # Configuración del agente IA
│   ├── rules/              # Reglas del agente
│   └── skills/            # Habilidades del agente
├── docs/
│   ├── PRD/                # Documentos de Requisitos
│   ├── specs/
│   │   ├── active/         # Specs en desarrollo
│   │   └── completed/       # Specs validadas
│   ├── plans/
│   │   ├── active/         # Planes de actuación en desarrollo
│   │   └── completed/      # Planes completados
│   ├── ERD/                # Diagramas de datos
│   ├── DISCOVERY_MAP.md    # Mapa del sistema
│   ├── DEUDA_TECNICA.md    # Deuda técnica
│   └── Diario_Sesiones.md   # Registro de sesiones
├── src/
│   ├── clases/             # Clases VBA
│   ├── modulos/            # Módulos estándar
│   └── formularios/        # Formularios
├── references/
│   └── Estructura_Datos.md
├── .engram/                # Memoria persistente
├── CHANGELOG.md
└── project_context.md       # Contexto del proyecto
```

### Paso 2: Configurar el contexto

Editar `project_context.md` con los datos del proyecto:

```markdown
# MiProyecto — Project Context

| Elemento | Valor |
|----------|-------|
| Lenguaje | VBA |
| Entorno | Microsoft Access |
| Frontend | MiApp.accdb |
| Backend | MiApp_Datos.accdb |
```

### Paso 3: Conectar con Access

Para sincronizar código VBA con archivos versionables:

```powershell
# Iniciar modo watch (edita en VS Code, se actualiza en Access)
node skills/access-vba-sync/cli.js watch --access "C:\Ruta\MiApp.accdb"

# Generar ERD del backend
node skills/access-vba-sync/cli.js generate-erd --backend "C:\Ruta\MiApp_Datos.accdb"
```

---

## Flujos de trabajo

### SDD — Spec-Driven Development

El flujo principal de desarrollo sigue el protocolo **SDD** con 4 fases y 2 STOPs:

```
HISTORIA DE USUARIO
        ↓
   FASE 1: Análisis → Spec
        ↓
   STOP 1: Aprobación de Spec
        ↓
   FASE 2: Implementación
        ↓
   STOP 2: Validación en Access
        ↓
   FASE 3: Iteración (si hay gaps)
        ↓
   FASE 4: Cierre → Archivar
```

---

## Cómo usa el framework el usuario

### Ejemplo 1: Nueva funcionalidad

**El usuario dice:**

> "Yo como miembro de calidad quiero que cuando registre una inspección, el sistema me avise automáticamente si el expediente ya tiene una inspección cerrada en el último mes, para evitar duplicidades."

**Qué sucede:**

1. **FASE 1 — Análisis (IA)**
   - Busca en memoria (Engram) si hay contexto previo.
   - Lee el DISCOVERY_MAP para localizar módulos afectados.
   - Lee el PRD del módulo de inspecciones.
   - Analiza el código fuente si es necesario.
   - **Detecta si es una Epic** (si afecta >3 módulos).

2. **Genera Spec**
   - Crea `docs/specs/active/spec-042-evitar-duplicidades/Spec-042_EvitarDuplicidades.md`
   - Incluye: historia de usuario, análisis de impacto, intervenciones, criterios de verificación.

3. **STOP 1 — Presentación al usuario**
   ```
   Spec-042: Evitar duplicidades en inspecciones
   
   Resumen: Añadir validación en GuardarInspeccion que compruebe si existe
   un registro en los últimos 30 días.
   
   Módulos afectados:
   - src/clases/InspeccionServicio.cls
   - src/modulos/InspeccionRepository.bas
   
   ¿Aprobado? Si/no → modificaciones
   ```

4. **FASE 2 — Implementación (IA)**
   - Implementa cada intervención.
   - Auto-verifica contra los criterios de la Spec.
   - Genera **Informe de Cambios UI** si hay cambios en formularios.

5. **STOP 2 — Validación en Access**
   - La IA **se detiene** y presenta los módulos modificados.
   - El usuario copia el código a su Access manualmente.
   - Compila y prueba.
   - Responde: `VALIDADO EN ACCESS: Spec-042` o describe los gaps.

6. **FASE 4 — Cierre (IA)**
   - Archiva la Spec en `completed/`
   - Actualiza el PRD
   - Actualiza DEUDA_TECNICA
   - Registra en Diario de Sesiones
   - Guarda en Engram
   - Genera CHANGELOG

---

### Ejemplo 2: Bug urgente en producción

**El usuario dice:**

> "Hay un bug crítico: cuando se elimina un expediente, los registros de inspecciones asociados quedan huérfanos. Hay que corregirlo ya."

**Qué sucede:**

1. El usuario ejecuta:
   ```powershell
   dysflow hotfix fix-huerfanos-inspeccion
   ```
   Esto crea la rama `hotfix-fix-huerfanos-inspeccion` desde `main`.

2. La IA genera una Spec con prefijo "hotfix".

3. Implementa la corrección.

4. Valida en Access.

5. **Cierre del hotfix**:
   ```powershell
   # El usuario ejecuta manualmente:
   git checkout main
   git pull
   git merge hotfix-fix-huerfanos-inspeccion
   dysflow next-release   # Obtiene el tag (ej: 2026-003)
   git tag 2026-003
   git push origin main
   git push origin 2026-003
   
   # Sincronizar develop
   git checkout develop
   git pull
   git merge main
   git push origin develop
   ```

---

### Ejemplo 3: Cambio de arquitectura

**El usuario dice:**

> "Queremos migrar el sistema de gestión de usuarios a un patrón repositorio limpio. Actualmente está todo en los formularios."

**Qué sucede:**

1. **Primero: RFC**
   - Se crea un RFC (Request for Comments) documentando el cambio de arquitectura.
   - El usuario aprueba el RFC.

2. **Después: Plan de Actuación**
   - La IA detecta que el cambio afecta >3 módulos.
   - Ofrece crear un **Plan de Actuación** con múltiples Specs coordinadas.

3. **Spec por Spec**:
   - Spec-043: Extraer RepositorioUsuario
   - Spec-044: Crear ServicioUsuario
   - Spec-045: Refactorizar formularios

4. Cada spec sigue el flujo SDD completo.

---

## Comandos CLI

| Comando | Descripción |
|---------|-------------|
| `dysflow spec <number>` | Crear rama `spec-{NNN}-{slug}` desde develop |
| `dysflow plan <number>` | Crear rama `plan-{NNN}-{slug}` para Planes de Actuación |
| `dysflow release` | Fusionar develop → main, crear tag YYYY-NNN |
| `dysflow hotfix <name>` | Crear rama hotfix desde main |
| `dysflow changelog <from>` | Generar changelog desde tag |
| `dysflow next-release` | Mostrar el próximo número de release |
| `dysflow init access` | Inicializar proyecto Access |
| `dysflow update` | Actualizar skills y rules |

---

## Skills del sistema

Las **skills** son capacidades que el agente de IA puede ejecutar:

| Skill | Función |
|-------|---------|
| `sdd-protocol` | Orquestador del flujo SDD |
| `spec-writer` | Generación de especificaciones técnicas |
| `prd-writer` | Generación y mantenimiento de PRDs |
| `hotfix` | Gestión de bugs urgentes |
| `rfc-writer` | Cambios de arquitectura |
| `access-vba-sync` | Sincronización Access ↔ código |
| `plan-writer` | Planes de actuación para Epics |
| `diario-sesion` | Registro de sesiones de desarrollo |

---

## Estructura de una Spec

```markdown
# Spec-042: [Título]

## Estado
🔵 ABIERTA | 🟡 EN PROGRESO | ✅ CERRADA

## 1. Resumen Técnico
- Problema: [qué falla o falta]
- Causa raíz: [por qué ocurre]
- Solución propuesta: [qué se va a hacer]

## 2. Historia de Usuario
> Como [rol], quiero [acción], para [beneficio].

## 3. Análisis de Impacto
- Módulos afectados
- Archivos a modificar
- Tablas de datos afectadas
- UI afectada

## 4. Plan de Intervención
- Intervención 1: [qué hacer]
- Intervención 2: [qué hacer]

## 5. Criterios de Verificación
- Auto-verificación (IA)
- Validación en Access (usuario)

## 6. Gaps y Decisiones
```

---

## Convenciones del proyecto

### Nomenclatura VBA

| Elemento | Convención | Ejemplo |
|----------|------------|---------|
| Clases/Módulos | PascalCase | `ClienteService.cls` |
| Funciones públicas | PascalCase | `GetClientesActivos()` |
| Variables locales | camelCase | `lngIdCliente` |
| Constantes | UPPER_SNAKE_CASE | `MAX_REINTENTOS` |
| Controles formulario | Prefijo tipo | `btnGuardar`, `txtNombre` |

### Nomenclatura Git

| Rama | Prefijo | Ejemplo |
|------|---------|---------|
| Feature | `spec-` | `spec-042-evitar-duplicidades` |
| Hotfix | `hotfix-` | `hotfix-fix-login` |
| Release | tag `YYYY-NNN` | `2026-003` |

---

## Memoria persistente (Engram)

El sistema usa **Engram** para recordar decisiones entre sesiones:

- **mem_search**: Buscar contexto previo.
- **mem_save**: Guardar decisiones importantes (arquitectura, bugs, aprendizajes).
- **mem_session_start**: Iniciar sesión.
- **mem_session_summary**: Cerrar sesión con resumen.

---

## Integración con Trae/OpenCode

1. Al iniciar una sesión, el agente ejecuta `mem_context`.
2. Detecta las skills disponibles en `.agent/skills/`.
3. Para cada tarea, sigue el protocolo SDD.
4. Al cerrar, ejecuta `mem_session_summary`.

---

## Casos de uso típicos

### Caso A: El usuario quiere una mejora

```
Usuario: "Quiero añadir un filtro por fecha en el listado de clientes."

IA: [Análisis] → [Genera Spec-XXX] → [STOP: espera aprobación]
Usuario: "Perfecto, adelante."
IA: [Implementa] → [STOP: espera validación en Access]
Usuario: "VALIDADO EN ACCESS: Spec-XXX"
IA: [Cierra: archiva, actualiza PRD, guarda en Engram, actualiza CHANGELOG]
```

### Caso B: El usuario reporta un bug

```
Usuario: "El cálculo de importe no funciona cuando el cliente tiene descuento."

IA: [Análisis del bug] → [Genera Spec con prefijo bug] → [STOP]
Usuario: "Adelante."
IA: [Implementa fix] → [STOP: espera validación]
Usuario: "VALIDADO EN ACCESS: Spec-YYY"
IA: [Cierra]
```

### Caso C: El usuario tiene una duda

```
Usuario: "¿Cómo funciona el módulo de facturación?"

IA: [mem_search] → [Lee PRD de facturación] → [Responde con explicación]
```

---

## Resolución de problemas

### Error: "Access is being used by another process"

```powershell
# Cerrar Access completamente y volver a intentar
node skills/access-vba-sync/cli.js generate-erd --backend "C:\Ruta\MiBackend.accdb"
```

### Error: "No se encontró la rama develop"

```powershell
# Crear develop desde main si no existe
git checkout main
git checkout -b develop
git push -u origin develop
```

### La Spec no pasa el STOP 1

- Revisar el análisis de impacto.
- Aclarar las ambigüedades en "Gaps y Decisiones".
- Añadir más contexto en la historia de usuario.

---

## Siguientes pasos

1. **Inicializar un proyecto**: `dysflow init access`
2. **Leer el project_context.md** y personalizarlo.
3. **Probar el flujo SDD** con una historia de usuario pequeña.
4. **Configurar el agente IA** en `.agent/AGENTS.md`.

---

## Recursos adicionales

- [Documentación de access-vba-sync](skills/access-vba-sync/README.md)
- [Plantilla de Spec](skills/spec-writer/references/SPEC-TEMPLATE.md)
- [Plantilla de PRD](skills/prd-writer/references/prd_template.md)
- [Project Context de referencia](project_context.md)
