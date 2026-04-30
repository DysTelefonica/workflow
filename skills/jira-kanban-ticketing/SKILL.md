---
name: jira-kanban-ticketing
description: >
  Redacta tickets Jira claros y operables para equipos Kanban: épicas, historias,
  tareas y tickets transversales con título normalizado, criterios de aceptación,
  validación y partición correcta del trabajo. Usar cuando haya que crear o mejorar
  issues de Jira para proyectos reales o para PORTFOLIO.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "1.0"
---

# jira-kanban-ticketing

## Cuándo usar

Usá esta skill cuando haya que:
- redactar un ticket nuevo en Jira
- mejorar un ticket vago o mal escrito
- decidir si algo es épica, historia, tarea o subtarea
- dividir una iniciativa grande en tickets movibles
- escribir criterios de aceptación y validación
- crear tickets transversales para `PORTFOLIO`

---

## Regla principal

Un ticket en Jira NO es un recordatorio personal.

Tiene que dejar claro:
- qué problema existe o qué cambio se quiere
- por qué importa
- qué significa terminarlo
- cómo se valida

Si no responde eso, el ticket está mal escrito.

---

## Tipos de issue y uso correcto

| Tipo | Cuándo usarlo | Qué debe contener |
|---|---|---|
| Epic | Iniciativa grande dentro de un proyecto | objetivo, alcance, entregables grandes |
| Historia | Entregable funcional o bloque documental con valor claro | resultado visible o valor concreto |
| Tarea | trabajo técnico, bug, setup, migración, análisis, chore | acción ejecutable y verificable |
| Subtarea | despiece puntual de trabajo ya acotado | paso operativo, no iniciativa |

### Regla crítica
Si el trabajo puede cerrarse sin necesidad de crear varios tickets subordinados, probablemente NO es una épica.

---

## Convención de títulos

Formato obligatorio:

```text
[TIPO] Descripción breve
```

### Prefijos permitidos
- `[BUG]`
- `[FEATURE]`
- `[TECH]`
- `[DOCS]`
- `[CHORE]`
- `[DECISION]`
- `[BLOCKER]`

### Buenos ejemplos
- `[BUG] El cierre del expediente no sincroniza NCs`
- `[TECH] Extraer validación de pendientes a helper reutilizable`
- `[DOCS] Documentar flujo de alta y edición de expedientes`
- `[BLOCKER] Falta backend local para validar solicitudes HPS`

### Malos ejemplos
- `revisar jira`
- `tema expedientes`
- `cosas pendientes`
- `mejoras`

### Regla de oro
Si el título no permite entender el trabajo sin abrir el ticket, está mal.

---

## Estructura mínima de un buen ticket

### 1) Descripción
Decir qué pasa hoy o qué se quiere cambiar.

### 2) Estado actual
Qué problema, restricción o vacío existe ahora.

### 3) Estado esperado
Qué comportamiento o resultado se espera al terminar.

### 4) Criterios de aceptación
Checklist verificable. No deseos vagos.

### 5) Notas técnicas
Solo lo necesario para orientar implementación.

### 6) Validación
Cómo confirmar que quedó bien.

---

## Plantilla base

Usar `assets/base-ticket-template.md`.

Estructura:
- Descripción
- Estado actual
- Estado esperado
- Criterios de aceptación
- Notas técnicas
- Validación

---

## Cómo escribir criterios de aceptación

### Bien
- [ ] al cerrar el formulario se persisten los cambios de cabecera
- [ ] `HayDatosPendientes` no marca pendiente permanente cuando el valor fue precargado
- [ ] queda disponible un checklist manual de validación en el ticket

### Mal
- [ ] funciona bien
- [ ] revisar que no falle
- [ ] dejarlo correcto

### Regla
Cada criterio tiene que poder responderse con sí/no.

---

## Cómo dividir trabajo grande

Si una iniciativa mezcla varias naturalezas, separarla.

### Separaciones típicas
- implementación
- documentación
- validación manual
- migración
- decisión de arquitectura

### Ejemplo
Mal:
- un ticket que mezcla fix de código + checklist manual + documentación + limpieza técnica

Bien:
- `[BUG] ...` para el fix
- `[DOCS] ...` para la documentación
- `[TECH] ...` para refactor o deuda
- `[CHORE] ...` para cleanup auxiliar

---

## PORTFOLIO vs proyecto real

Usar `PORTFOLIO` si el ticket trata de:
- prioridades entre proyectos
- gobierno transversal
- estándares
- bloqueos compartidos
- decisiones que afectan a varios sistemas

Usar un proyecto real si el ticket trata de:
- bug
- feature
- refactor
- documentación propia
- validación o operación propia del sistema

### Regla
No mandar a `PORTFOLIO` lo que en realidad es trabajo interno de un proyecto.

---

## Épicas bien escritas

Usar `assets/epic-template.md`.

Una épica debe incluir:
- objetivo
- alcance
- fuera de alcance
- riesgos
- entregables grandes
- sugerencia de tickets hijos

### Regla crítica
La épica no debe reemplazar el backlog detallado. Debe servir para descomponerlo.

---

## Tickets de documentación

Los `[DOCS]` también deben ser operables.

Mal:
- `[DOCS] Documentación de Expedientes`

Bien:
- `[DOCS] Documentar flujo de guardado automático en Expedientes`
- `[DOCS] Crear checklist manual de validación de cierre de expediente`

### Regla
La documentación debe tener un alcance concreto, no una etiqueta genérica.

---

## Tickets técnicos

Para `[TECH]`, incluir siempre al menos uno de estos focos:
- deuda técnica
- desacoplamiento
- simplificación
- robustez
- mantenibilidad
- soporte a validación futura

### Regla
No disfrazar trabajo funcional como técnico ni al revés.

---

## Prompts base para agentes

### Crear un ticket bien redactado

```text
Redactá este ticket Jira para un flujo Kanban.
No uses lenguaje Scrum.
Usá un título con prefijo correcto.
Incluí: descripción, estado actual, estado esperado, criterios de aceptación verificables, notas técnicas y validación.
Si el trabajo es demasiado grande o mezcla varias naturalezas, proponé dividirlo.
```

### Mejorar un ticket vago

```text
Reescribí este ticket Jira para que quede operable.
Eliminá ambigüedad.
Convertí deseos vagos en criterios verificables.
Si falta contexto mínimo, explicá qué falta.
Si mezcla varios trabajos, proponé separación.
```

---

## Comandos útiles

### Instalar esta skill en OpenCode

```powershell
Copy-Item -Recurse -Force "C:\Proyectos\workflow\skills\jira-kanban-ticketing" "$env:USERPROFILE\.config\opencode\skills\"
```

---

## Checklist final para la IA

- [ ] ¿Elegí el tipo de issue correcto?
- [ ] ¿El título usa prefijo y dominio claro?
- [ ] ¿El ticket explica problema o cambio con precisión?
- [ ] ¿Los criterios de aceptación son verificables?
- [ ] ¿La validación es concreta?
- [ ] ¿Separé trabajo grande o mezclado si hacía falta?
- [ ] ¿Mandé el ticket al espacio/proyecto correcto?
