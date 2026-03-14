# RFC-{NNN}: {Título del cambio de arquitectura}

**Estado:** `Borrador` | `En revisión` | `Aprobado` | `Rechazado` | `Implementado`
**Fecha:** {YYYY-MM-DD}
**Autor:** {nombre o rol}
**Specs relacionadas:** Spec-XXX, Spec-YYY (si las hay)

---

## 1. Problema

> Describe el problema o limitación arquitectónica que motiva este RFC.
> ¿Por qué la arquitectura actual no es suficiente?

{descripción del problema}

## 2. Contexto

> Información de fondo necesaria para entender el problema.
> Módulos afectados, dependencias actuales, decisiones previas relevantes.

{contexto}

## 3. Propuesta

> Descripción concreta del cambio arquitectónico propuesto.
> Sé específico: qué cambia, cómo funciona después del cambio.

{propuesta detallada}

## 4. Alternativas consideradas

> ¿Qué otras opciones se evaluaron? ¿Por qué se descartaron?

| Alternativa | Motivo de descarte |
| :--- | :--- |
| {alternativa 1} | {motivo} |
| {alternativa 2} | {motivo} |

## 5. Impacto

### Módulos afectados

| Módulo / Archivo | Tipo de cambio | Notas |
| :--- | :--- | :--- |
| {archivo.bas} | Refactor / Nuevo / Eliminado | {notas} |

### Cambios en modelo de datos

> ¿Este RFC modifica tablas, campos o relaciones en la base de datos?

- [ ] No aplica
- [ ] Sí — descripción: {detalles}

### Cambios en UI (formularios)

> ¿Este RFC modifica formularios visibles al usuario?

- [ ] No aplica
- [ ] Sí — descripción: {detalles}

### Riesgos

| Riesgo | Probabilidad | Mitigación |
| :--- | :--- | :--- |
| {riesgo 1} | Alta / Media / Baja | {mitigación} |

## 6. Plan de implementación

> ¿Este RFC se implementa en una sola Spec o en múltiples?
> Si son múltiples, ¿en qué orden y con qué dependencias?

- [ ] Una sola Spec: Spec-{NNN}
- [ ] Múltiples Specs (Plan de Actuación):
  1. Spec-{NNN}: {descripción}
  2. Spec-{NNN+1}: {descripción} _(depende de Spec-{NNN})_

## 7. Criterio de aceptación

> ¿Cómo se verifica que el RFC está correctamente implementado?

- [ ] {criterio 1}
- [ ] {criterio 2}

## 8. Decisión

> Completar cuando el RFC sea aprobado o rechazado.

**Decisión:** {Aprobado / Rechazado}
**Fecha:** {YYYY-MM-DD}
**Justificación:** {motivo de la decisión}

---

> **Cómo usar esta plantilla**
> 1. Crea el archivo en `docs/RFC/RFC-{NNN}_{slug}.md`
> 2. Completa las secciones — omite solo las que genuinamente no aplican (indica "N/A")
> 3. Comparte con el equipo para revisión antes de generar las Specs
> 4. Una vez aprobado, actualiza el estado y crea las Specs correspondientes
