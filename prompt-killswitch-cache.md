# Prompt para incorporar kill-switch de cache

Quiero anadir una capacidad de "modo seguro" para cache (kill-switch) por seguridad operativa.

No implementes codigo aun; actualiza documentacion (RFC + Plan + Spec-008 y la spec que corresponda) con esta mejora.

## Objetivo
Poder desactivar la cache en produccion de forma inmediata si algo va mal, mantener la aplicacion operativa sin cache, investigar, y reactivarla despues.

## Cambios obligatorios

1) RFC-001 (`docs/rfcs/RFC-001_arquitectura-cache-viewmodel.md`)
- Anadir seccion "Kill-switch de cache (modo seguro)".
- Definir:
  - Flag global `CacheEnabled` (Boolean) en configuracion persistida.
  - Comportamiento cuando `CacheEnabled=False`:
    - no leer cache,
    - no escribir/reconstruir cache,
    - usar ruta directa de datos (sin cache).
  - Activacion/desactivacion sin despliegue (desde funcion administrativa / inmediato).
  - Rehabilitacion controlada (activar + opcional precalentado manual).
- Riesgos/mitigaciones:
  - degradacion de rendimiento temporal,
  - consistencia funcional garantizada.

2) PLAN-002 (`docs/plans/active/plan-002-cache-viewmodel-rfc001/PLAN_002_Cache_ViewModel_RFC001.md`)
- Anadir tarea nueva (T-10) "Kill-switch operativo de cache".
- Dependencias: despues de T-03, T-06, T-08 (y compatible con T-09 precalentado).
- Criterios de aceptacion de T-10:
  - con flag OFF la app funciona completa sin cache,
  - con flag ON vuelve a usar cache,
  - cambio de estado no requiere despliegue,
  - logging de cuando/quien activo/desactivo.

3) Specs
- Crear nueva spec:
  - `docs/specs/active/spec-010-killswitch-cache/Spec-010_KillSwitch_Cache.md`
- En Spec-008 anadir referencia cruzada a Spec-010.
- En Spec-010 incluir:
  - diseno del flag y punto unico de lectura (`IsCacheEnabled()`),
  - comportamiento en lectura/escritura de cache con OFF/ON,
  - pruebas en Ventana Inmediato:
    - `? CacheConfig_SetEnabled(False)` -> flujo sin cache
    - `? CacheConfig_SetEnabled(True)` -> flujo con cache
  - criterio de seguridad:
    - "si OFF, ninguna operacion de cache puede ejecutarse"
  - rollback:
    - volver a OFF de forma inmediata.

4) Mantener decisiones actuales
- sin TTL en detalle
- refresco manual
- coherencia cascada AR -> AC -> NC
- atomicidad CRUD + operacion minima de cache (si cache esta habilitada)

## Salida obligatoria
- rutas de archivos modificados/creados
- resumen de cambios por archivo
- tabla de aceptacion T-10 / Spec-010 (ON vs OFF)
- confirmacion final:
  "Existe kill-switch de cache para desactivar y reactivar el sistema sin despliegue"
