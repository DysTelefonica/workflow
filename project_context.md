# {NOMBRE_PROYECTO} — Project Context

> Este archivo define el **contexto técnico completo del proyecto**.
> Es leído por varias skills del sistema (especialmente `prd-writer` y `spec-writer`)
> para generar documentación técnica consistente sin rutas ni convenciones hardcodeadas.

Este documento describe:

- estructura del repositorio
- stack tecnológico
- arquitectura
- convenciones del proyecto
- workflow de desarrollo

---

# 1. Estructura del repositorio

| Recurso | Ruta |
|-------|------|
| Código fuente — formularios | `src/formularios/` |
| Código fuente — clases | `src/clases/` |
| Código fuente — módulos | `src/modulos/` |
| PRDs | `docs/PRD/` |
| Specs activas | `docs/specs/active/` |
| Specs completadas | `docs/specs/completed/` |
| DISCOVERY_MAP | `docs/DISCOVERY_MAP.md` |
| DEUDA_TECNICA | `docs/DEUDA_TECNICA.md` |
| Diario de sesiones | `docs/Diario_Sesiones.md` |
| Modelo de datos | `references/Estructura_Datos.md` |

---

# 2. Infraestructura del agente

El proyecto utiliza un sistema de desarrollo asistido por IA basado en **skills y memoria persistente**.

| Recurso | Ruta |
|-------|------|
| Configuración del agente | `.agent/AGENTS.md` |
| Reglas globales del agente | `.agent/rules/` |
| Skills del agente | `.agent/skills/` |
| Memoria Engram | `.agent/engram/` |

Las skills principales que interactúan con este proyecto son:

- `sdd-protocol`
- `prd-writer`
- `spec-writer`
- `access-vba-sync`
- `hotfix`
- `rfc-writer`

---

# 3. Stack tecnológico

| Elemento | Valor |
|-------|------|
| Lenguaje | VBA |
| Entorno | Microsoft Access |
| Motor de datos | DAO |
| Frontend | Formularios Access |
| Backend | Base de datos Access |

Bases de datos del proyecto:

| Tipo | Archivo |
|-----|------|
| Frontend | `{nombre}.accdb` |
| Backend principal | `{nombre_backend}.accdb` |
| Backend externos | `{nombre_externa}.accdb` |

---

# 4. Arquitectura del proyecto

El proyecto sigue una arquitectura **en tres capas**.

| Capa | Ubicación | Responsabilidad |
|----|------|------|
| UI | `src/formularios/` | Formularios Access, interacción con usuario |
| Negocio | `src/clases/` | Lógica de negocio, validaciones |
| Datos | `src/modulos/` | Acceso a datos DAO y SQL |

### Flujo típico
Formulario
→ instancia clase de negocio
→ ejecuta método de negocio
→ módulo de datos
→ consulta SQL

---

# 5. Convenciones de nomenclatura

## Tablas

| Tipo | Patrón | Ejemplo |
|----|----|----|
| Tabla principal | `Tb` + PascalCase | `TbEventos` |
| Tabla auxiliar | `TbAux` + PascalCase | `TbAuxEventos` |
| Identificadores | `ID` + Entidad | `IDEvento` |

---

## Código VBA

| Elemento | Convención |
|------|------|
| Clases | PascalCase |
| Módulos | PascalCase |
| Formularios | `Form_` + nombre |
| Funciones públicas | PascalCase |
| Variables locales | camelCase |
| Constantes | UPPER_SNAKE_CASE |

---

## Controles de formulario

| Prefijo | Tipo |
|------|------|
| `btn` | botón |
| `txt` | textbox |
| `cmb` | combo |
| `lst` | lista |
| `lbl` | label |
| `frm` | subformulario |

---

# 6. Tipos de datos Access

Para documentar PRDs se utiliza la siguiente traducción:

| Código | Tipo |
|------|------|
| 1 | Boolean |
| 3 | Integer |
| 4 | Long |
| 6 | Single |
| 7 | Double |
| 8 | Date/Time |
| 10 | Text |
| 12 | Memo |

Nunca documentar el código numérico en PRDs.

---

# 7. Gestión de errores

| Tipo | Implementación |
|----|----|
| Errores de negocio | `Err.Raise 1000+` |
| Mensajes usuario | `MsgBox` |
| Logging | `Debug.Print` |
| Propagación error | `Optional ByRef p_Error As String` |

---

# 8. Integraciones externas

| Sistema | Descripción | Punto de entrada |
|------|------|------|
| {sistema} | {descripción} | {método o variable} |

---

# 9. Workflow de desarrollo (SDD)

El proyecto utiliza **Specification Driven Development**.

Flujo:
Historia de usuario
↓
Spec
↓
STOP aprobación
↓
Implementación
↓
Validación

Reglas:

- Nunca escribir código sin Spec aprobada.
- Toda modificación debe tener una Spec.

---

# 10. Workflow Git

El proyecto sigue un **Git Flow simplificado**.

| Rama | Uso |
|----|----|
| main | código en producción |
| develop | integración de features |
| spec branches | implementación de specs |

Convención de ramas:
spec-XXX-descripcion
Ejemplo:
spec-042-fix-calculo-importes

---

# 11. Sistema de releases

Las releases se etiquetan con el formato:
YYYY-NNN

Ejemplo:
2026-003
Hotfixes:
YYYY-NNN.1
Ejemplo:
2026-003.1

---

# 12. Deployment en producción

Arquitectura de ejecución:
Servidor
├─ Backend.accdb
└─ recursos/
└─ Aplicacion.accde


Los usuarios ejecutan la aplicación desde **copias locales**.

Directorio local típico:


%APPDATA%\Aplicaciones_Dys{Aplicacion}


Flujo:

1. Lanzadera verifica versión.
2. Si hay versión nueva → copia archivos desde servidor.
3. Usuario ejecuta frontend local.

Esto evita bloqueos de archivos compartidos.

---

# 13. Rollback

Las versiones anteriores del frontend se almacenan en:


versions/


Si una release falla:

1. se recompila ACCDE de una versión anterior
2. se publica como nueva versión
3. la lanzadera obliga actualización.

---

# 14. Notas para generación de PRDs

Al generar PRDs el agente debe:

- usar las rutas de este documento
- respetar convenciones de nombres
- documentar solo campos reales del modelo de datos
- evitar duplicar información existente

---

# 15. Notas de dominio del proyecto

Añadir aquí cualquier conocimiento específico del dominio.

Ejemplos:

- reglas de negocio
- restricciones de datos
- convenciones históricas del proyecto