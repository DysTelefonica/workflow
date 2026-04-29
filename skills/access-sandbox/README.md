# access-sandbox

Provisiona backends Access locales para trabajo offline. Copia backends de producción al sandbox local y reescribe las tablas vinculadas para que apunten a local.

## Qué hace

- Copia todos los backends de una carpeta de origen al sandbox
- Reescribe `TableDef.Connect` de cada tabla vinculada para apuntar al sandbox local
- Elimina tablas vinculadas si el backend destino no existe o `RefreshLink()` falla — **seguridad**: nunca queda referencia a producción
- Crea zip de seguridad antes de tocar nada

## Quick start

```powershell
# 1. Establecer password (obligatorio)
$env:ACCESS_SANDBOX_PW = 'dpddpd'

# 2. Copiar y adaptar configs/example.dev.json
#    (definir sourceFolder y localSandboxPath)

# 3. Ejecutar
& "scripts/sync-backends.ps1" -ConfigPath "configs/example.dev.json"
```

## Requisitos

- Windows + Access Database Engine (DAO.DBEngine.120)
- PowerShell 5.1+
- `ACCESS_SANDBOX_PW` obligatoria

## Estructura

| Ruta | Descripción |
|---|---|
| `scripts/sync-backends.ps1` | Motor de sync |
| `scripts/ConvertLinkedAccessTablesToLocal.ps1` | Convierte frontend a local |
| `configs/example.dev.json` | Template — copiar y adaptar |
