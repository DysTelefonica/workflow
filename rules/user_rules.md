# Reglas de Conducta del Agente

Estas reglas se aplican en todas las sesiones, en todos los proyectos,
independientemente del stack, dominio o fase del proyecto.

---

## 1. Orden de consulta obligatorio

Antes de abrir cualquier archivo o escribir cualquier código:

```
1. mem_context / mem_search     ← SIEMPRE primero, sin excepción
2. docs/PRD/                    ← solo si Engram no cubre
3. docs/DISCOVERY_MAP.md        ← solo si el PRD no cubre
4. src/                         ← último recurso, solo lo imprescindible
```

Saltar este orden desperdicia contexto y genera inconsistencias con decisiones previas.

---

## 2. Prerequisito de PRD

**Sin PRD del módulo afectado, no hay Spec. Sin Spec, no hay código.**

Si el módulo afectado no tiene PRD → activar `prd-writer` antes de cualquier otra acción.
No hay excepciones, ni siquiera para "cambios pequeños" o "fixes evidentes".

---

## 3. Inicio de sesión obligatorio

Siempre, sin excepción, al comenzar cualquier sesión:

```
1. mem_session_start
2. mem_context
```

Si Engram devuelve una Spec activa → retomarla sin pedir al usuario que repita el contexto.

---

## 4. Cierre de sesión obligatorio

Siempre, sin excepción, al finalizar cualquier sesión:

```
mem_session_summary:
  Goal:         qué se quería lograr
  Discoveries:  hallazgos de arquitectura, bugs, patrones, FKs
  Accomplished: qué quedó completado y validado
  Files:        archivos creados o modificados (rutas relativas)

mem_session_end
```

**Sin `mem_session_summary`, la próxima sesión empieza ciega.**
Si el contexto se compacta antes de ejecutarlo → llamar a `mem_context` inmediatamente
al retomar y continuar desde el punto de interrupción.

---

## 5. Calidad de memoria Engram

Antes de cualquier `mem_save`, aplicar la rule `engram-memory-quality.md`.
Un save de calidad vale más que diez de ruido.

---

## 6. Validación humana antes de implementar

Nunca implementar código sin que el usuario haya aprobado la Spec.
El STOP 1 del sdd-protocol es obligatorio e innegociable.

---

## 7. Entrega manual de código

El agente entrega listas de módulos modificados.
El usuario los importa manualmente en Access/VBA.
El agente nunca ejecuta importaciones ni compilaciones.

---

## 8. Zero regresiones

Un fix o una nueva funcionalidad no deben romper lo que ya funcionaba.
Antes de implementar: identificar módulos adyacentes en riesgo.
Después de implementar: verificar que los criterios de la Spec se cumplen sin excepción.

---

## 9. No sobreingeniería

Resolver exactamente lo que pide la historia de usuario.
Ni más, ni menos.
Si durante la implementación aparece algo que parece necesario pero no está en la Spec
→ documentarlo como gap o deuda técnica, no implementarlo sin aprobación.

---

## 10. Ante compactación de contexto

Si el agente detecta que el contexto se ha compactado (empieza "en blanco"):

```
1. mem_session_start
2. mem_context
3. mem_search "[proyecto] spec activa"
```

Recuperar el estado y continuar sin pedir al usuario que repita lo que ya se discutió.

---

## 11. Ante feedback de calidad

Si el usuario indica que un PRD, Spec o cualquier entregable no cumple el estándar:

```
1. NO continuar al siguiente módulo
2. Releer la skill correspondiente
3. Releer engram-memory-quality.md
4. Corregir el entregable punto por punto
5. Presentar exactamente qué se corrigió
6. Solo entonces continuar
```

---

## 12. No crear placeholders vacíos

No crear archivos de documentación hasta tener contenido real para incluir en ellos.
`DISCOVERY_MAP.md`, `DEUDA_TECNICA.md`, PRDs — se crean cuando hay datos reales.
Las estructuras vacías no ayudan y confunden en la próxima sesión.
