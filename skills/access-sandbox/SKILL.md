---
name: access-sandbox
description: Access local sandbox provisioning and VBA synchronization. Use when working with Access backends in local mode, converting linked tables to local, synchronizing production backends to a sandbox, or managing Access VBA code in local development environments.
---

# Access Sandbox

Provisiona backends Access locales y sincroniza código VBA entre archivos Access y el sistema de archivos.

## Scripts

### sync-backends.ps1

Sincroniza backends Access de producción hacia un sandbox local, reescribiendo las tablas vinculadas para que funcionen offline.

**Ubicación:** `scripts/sync-backends.ps1`

**Uso:**
```powershell
$env:ACCESS_SANDBOX_PW = 'dpddpd'
& "scripts/sync-backends.ps1" -ConfigPath "configs/example.dev.json"
```

**Flujo en tres pasos:**
1. **Zip de seguridad** — comprime lo que haya en el sandbox (conserva `.zip` existentes)
2. **Limpieza** — borra todo menos los `.zip`
3. **Copia + Revinculación two-pass:**
   - Pass 1: abre cada backend, detecta tablas vinculadas a otros `.accdb`, construye un plan
   - Pass 2: aplica `TableDef.Connect = local + RefreshLink()`. Si falla, **elimina la tabla** (nunca deja referencia a producción)

**Config JSON:**
```json
{
  "sourceFolder": "C:\\...\\000datoslocal",
  "localSandboxPath": "C:\\00repos\\datos",
  "productionPaths": [...]
}
```

> **Password obligatoria:** `ACCESS_SANDBOX_PW`. Sin ella el script aborta.

**Seguridad:**
- Si `RefreshLink()` falla, la tabla vinculada se **elimina** — no queda jamás ninguna referencia a `\\datoste\...`
- Si el backend destino no existe, la tabla vinculada se **elimina**
- Verificación final garantiza cero `\\datoste` en todos los archivos

**Resultado típico:**
- 99 vínculos reescritos OK
- 4 eliminados (backends no existentes en destino o tablas pre-rotas en origen)

### ConvertLinkedAccessTablesToLocal.ps1

Convierte tablas vinculadas de un frontend Access en tablas locales reales. El frontend queda autocontenido, sin dependencias externas.

**Ubicación:** `scripts/ConvertLinkedAccessTablesToLocal.ps1`

## Estructura del skill

```
access-sandbox/
├── SKILL.md                                    # Este archivo
├── README.md                                   # Guía de uso rápido
├── scripts/
│   ├── sync-backends.ps1                       # Sync backends producción → sandbox
│   └── ConvertLinkedAccessTablesToLocal.ps1    # Convertir frontend a local
└── configs/
    ├── backends_config.json                   # Configuración de producción (NO commitear)
    └── example.dev.json                      # Template de configuración — copiar y adaptar
```

## Relación entre skills

| Escenario | Skill a usar |
|---|---|
| Sincronizar backends de producción al sandbox local | `sync-backends.ps1` |
| Convertir un frontend con vínculos a tablas locales | `ConvertLinkedAccessTablesToLocal.ps1` |
| Consultar o modificar datos en un `.accdb` | `access-query` |

## Flujo combinado típico

```
1. sync-backends.ps1          → copia backends de \\datoste a C:\00repos\datos
                                 y reescribe los TableDef.Connect
2. ConvertLinkedAccessTablesToLocal.ps1  → convierte el frontend en autónomo
3. access-query                → opera sobre el resultado
```
