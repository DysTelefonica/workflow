# Runbook operativo - Kill-switch de cache

## Objetivo
Desactivar y reactivar la cache de forma segura en produccion sin despliegue, manteniendo la aplicacion operativa.

## Alcance
- Aplicacion No Conformidades
- Cache de gestion/listado
- Cache de detalle (NC + hijos)

## Precondiciones
- Tener permisos de operacion sobre la aplicacion/BD.
- Confirmar que existen funciones de control de cache en Ventana Inmediato.
- Informar a usuarios clave de posible degradacion temporal de rendimiento al desactivar cache.

## Comandos de referencia (Ventana Inmediato)
> Ajustar nombres si en implementacion final cambian las funciones.

```vba
' Estado actual
? CacheConfig_IsEnabled()

' Desactivar cache (modo seguro)
? CacheConfig_SetEnabled(False)

' Reactivar cache
? CacheConfig_SetEnabled(True)

' (Opcional) Precalentar cache tras reactivacion
? CacheNCProyecto_PrecalentarCompleta()
```

## Procedimiento A - Activar modo seguro (cache OFF)
1. Confirmar incidencia (desalineacion, error de cache, comportamiento anomalo).
2. Registrar hora, usuario operador y motivo.
3. Ejecutar en Inmediato: `? CacheConfig_SetEnabled(False)`.
4. Verificar estado: `? CacheConfig_IsEnabled()` debe devolver `False`.
5. Prueba funcional minima:
   - abrir gestion NC,
   - abrir detalle NC,
   - guardar un cambio sencillo.
6. Comunicar a usuarios: sistema operativo, rendimiento temporalmente menor.

## Procedimiento B - Reactivar cache (cache ON)
1. Confirmar que la causa raiz esta mitigada.
2. Ejecutar en Inmediato: `? CacheConfig_SetEnabled(True)`.
3. Verificar estado: `? CacheConfig_IsEnabled()` debe devolver `True`.
4. (Recomendado) Ejecutar precalentado: `? CacheNCProyecto_PrecalentarCompleta()`.
5. Validar:
   - gestion con filtros,
   - detalle NC,
   - CRUD simple con invalidacion esperada.
6. Comunicar restablecimiento de cache.

## Validaciones de aceptacion operativa
- Con cache OFF, la app funciona end-to-end sin errores funcionales.
- Con cache ON, se recupera mejora de rendimiento esperada.
- No hay desalineacion visible tras CRUD y refresco manual.

## Criterios de rollback
- Si al reactivar cache aparecen errores o desalineacion:
  1. Volver a OFF inmediatamente: `? CacheConfig_SetEnabled(False)`.
  2. Mantener operacion sin cache.
  3. Abrir analisis tecnico antes de nuevo intento.

## Evidencias a guardar
- Fecha/hora de OFF y ON.
- Operador que ejecuta cambios.
- Resultado de `CacheConfig_IsEnabled()`.
- Resultado de pruebas minimas (gestion/detalle/guardar).

## Plantilla de comunicacion interna
"Se activa modo seguro de cache por incidencia. La aplicacion sigue operativa; puede haber menor rendimiento temporal. Se notificara restablecimiento cuando la cache se reactive y valide."
