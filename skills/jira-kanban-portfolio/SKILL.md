---
name: jira-kanban-portfolio
description: >
  Define y opera Jira con enfoque Kanban por dominios reales: un espacio/proyecto por
  sistema, PORTFOLIO como capa de gobierno, épicas para iniciativas grandes y tareas/
  historias como unidades movibles. Usar cuando haya que crear espacios, tableros,
  épicas, tareas, convenciones o documentación operativa en Jira.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "1.0"
---

# jira-kanban-portfolio

## Cuándo usar

Usá esta skill cuando el trabajo implique cualquiera de estos contextos:
- crear o reorganizar Jira para proyectos reales
- decidir si algo va en `PORTFOLIO` o en un proyecto operativo
- crear épicas, historias o tareas en espacios como `EXPEDIENTES`, `WORKFLOW`, etc.
- definir convenciones de nombres, tipos de issue o criterio de terminado
- documentar backlog, deuda técnica, decisiones y bloqueos
- evitar que la IA derive a Scrum cuando el equipo trabaja en Kanban

---

## Modelo de trabajo obligatorio

### 1) Kanban, no Scrum

Regla dura:
- NO proponer sprints
- NO proponer sprint planning
- NO usar velocity ni story points por defecto
- NO hablar de sprint backlog salvo que el usuario lo pida explícitamente

Pensar siempre en:
- Backlog
- En curso
- Bloqueado
- Hecho

### 2) Un espacio/proyecto por dominio real

Ejemplos:
- `EXPEDIENTES`
- `WORKFLOW`
- `CONDOR`
- `NO_CONFORMIDADES`

No mezclar proyectos distintos en el mismo backlog operativo.

### 3) `PORTFOLIO` no es trabajo operativo normal

`PORTFOLIO` sirve para:
- gobierno transversal
- visión global
- prioridades entre proyectos
- bloqueos comunes
- decisiones que afectan a varios espacios
- estándares/documentación de trabajo

`PORTFOLIO` NO es el lugar para:
- bugs concretos de un proyecto
- tareas técnicas del día a día
- duplicar el tablero operativo de cada espacio

---

## Jerarquía recomendada

| Nivel | Uso correcto | No usar para |
|---|---|---|
| Espacio/Proyecto | Dominio real (`EXPEDIENTES`, `WORKFLOW`) | Mezclar varios dominios |
| Epic | Iniciativa grande dentro de un proyecto | Representar el proyecto entero si ya existe un espacio propio |
| Historia | Entregable funcional o bloque documental con valor claro | Trabajo técnico ambiguo |
| Tarea | Trabajo técnico, análisis, setup, bug, migración, chore | Meter alcance funcional gigante |
| Subtarea | Despiece operativo puntual | Crear pseudo-proyectos |

### Regla crítica
Si ya existe un espacio/proyecto real, la épica representa una **línea grande de trabajo dentro de ese proyecto**, NO el proyecto entero.

---

## Regla de decisión rápida

| Quiero crear... | Dónde va |
|---|---|
| prioridad transversal entre proyectos | `PORTFOLIO` |
| bloqueo común o dependencia entre sistemas | `PORTFOLIO` |
| bug concreto de Expedientes | `EXPEDIENTES` |
| refactor técnico de Workflow | `WORKFLOW` |
| documentación propia de un sistema | su espacio/proyecto |
| estándar transversal de Jira o docs | `PORTFOLIO` |

---

## Convención de títulos

Formato recomendado:

```text
[TIPO] Descripción breve
```

### Tipos permitidos
- `[BUG]`
- `[FEATURE]`
- `[TECH]`
- `[DOCS]`
- `[CHORE]`
- `[DECISION]`
- `[BLOCKER]`

### Ejemplos
- `[BUG] Autoguardado no sincroniza NCs al cerrar expediente`
- `[TECH] Separar helper de guardado automático por responsabilidades`
- `[DOCS] Documentar flujo de alta y edición de expediente`
- `[DECISION] Definir estructura Kanban por espacios reales`
- `[BLOCKER] Dependencia entre acceso a backend y validación manual`

### Regla de calidad
El título tiene que responder por sí solo:
- qué clase de trabajo es
- qué cambió o qué problema existe
- en qué lenguaje de dominio estamos

No usar títulos basura como:
- `arreglar cosas`
- `pendientes`
- `tema jira`
- `mejoras varias`

---

## Cómo partir trabajo grande

### Si toca varias capas, dividir
Si una iniciativa toca varias áreas, dividirla en tickets separados.

Ejemplos de partición útil:
- VBA / Access frontend
- backend / datos
- documentación
- migración
- validación manual

### Regla
No crear un ticket monstruo que mezcle:
- implementación
- validación
- documentación
- migración
- decisiones de arquitectura

Separar por unidad movible y verificable.

---

## Plantillas mínimas

### Epic
Usar `assets/epic-template.md`

Debe incluir:
- Objetivo
- Alcance
- Fuera de alcance
- Riesgos
- Checklist de grandes entregables

### Historia / Tarea
Usar `assets/task-template.md`

Debe incluir:
- Descripción
- Estado actual
- Estado esperado
- Criterios de aceptación
- Notas técnicas
- Validación

### Ticket transversal de portfolio
Usar `assets/portfolio-item-template.md`

Debe incluir:
- Contexto
- Impacto transversal
- Próximo paso
- Riesgos / bloqueos

---

## Criterios de aceptación: regla dura

Nunca crear tickets sin criterios verificables.

Mal:
- “dejarlo mejor”
- “revisar si funciona”

Bien:
- “al cerrar el formulario, los cambios de cabecera sincronizan NCs”
- “`HayDatosPendientes` no marca pendiente permanente cuando `NModificado` fue precargado”
- “queda documentado el flujo de validación manual en el ticket”

---

## Documentación dentro de Jira

No subir documentación como ruido suelto. Ordenarla así:

### En proyectos reales
- historias o tareas `[DOCS]` para:
  - contexto del sistema
  - arquitectura
  - backlog inicial
  - decisiones técnicas
  - checklist de validación

### En `PORTFOLIO`
- `[DECISION]` para normas globales
- `[DOCS]` para estándares transversales
- `[BLOCKER]` para bloqueos comunes
- `[CHORE]` para reorganizaciones globales

---

## Tableros: criterio correcto

### En cada proyecto real
Debe existir su tablero operativo principal.

Ejemplo:
- `EXPEDIENTES` → tablero operativo de Expedientes
- `WORKFLOW` → tablero operativo de Workflow

### En `PORTFOLIO`
Los tableros deben ser de visión o coordinación, por ejemplo:
- General
- Bloqueados
- Prioridades actuales
- Gobierno / documentación transversal

### Regla crítica
No duplicar el tablero operativo principal de un proyecto dentro de `PORTFOLIO` salvo que sea una vista agregada o un filtro específico.

---

## Flujo Kanban recomendado

Mínimo viable:
- Backlog
- En curso
- Bloqueado
- Hecho

Si hace falta más granularidad:
- Backlog
- Ready
- En curso
- En validación
- Bloqueado
- Hecho

### Regla
No complejizar el flujo antes de tener hábito real de uso.

---

## Prompts base para agentes

### 1) Crear trabajo en Jira sin caer en Scrum

```text
Trabajá este Jira con enfoque KANBAN, no Scrum.
El espacio/proyecto actual representa un dominio real.
No propongas sprints, sprint planning, velocity ni story points.
Si creás issues, usá títulos con prefijo ([BUG], [FEATURE], [TECH], [DOCS], [CHORE], [DECISION], [BLOCKER]).
Cada issue debe incluir criterios de aceptación verificables.
Si el trabajo toca varias capas, dividilo en tickets separados.
```

### 2) Decidir si algo va a `PORTFOLIO` o a un proyecto real

```text
Decidí el destino correcto del trabajo:
- `PORTFOLIO` si es transversal, de gobierno, prioridad global, bloqueo compartido o estándar.
- proyecto real si es implementación, bug, refactor, documentación o validación propia de ese sistema.
Explicá la decisión con una frase corta antes de crear el issue.
```

---

## Comandos útiles

### Instalar esta skill en OpenCode desde el repo local

```powershell
Copy-Item -Recurse -Force "C:\Proyectos\workflow\skills\jira-kanban-portfolio" "$env:USERPROFILE\.config\opencode\skills\"
```

### Verificar que OpenCode ya tiene Atlassian MCP

```powershell
opencode mcp list
opencode mcp debug atlassian
```

---

## Checklist final para la IA

- [ ] ¿Estoy pensando en Kanban y no en Scrum?
- [ ] ¿El issue va en el espacio/proyecto correcto?
- [ ] ¿`PORTFOLIO` se usa solo para coordinación/gobierno?
- [ ] ¿La épica representa una iniciativa y no el proyecto entero?
- [ ] ¿El ticket tiene título claro con prefijo?
- [ ] ¿Tiene criterios de aceptación verificables?
- [ ] ¿Separé trabajo multi-capa en tickets distintos si hacía falta?
- [ ] ¿La documentación quedó ordenada y no mezclada con barro?
