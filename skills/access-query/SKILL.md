---
name: access-query
description: >
  Ejecuta SQL seguro contra backends Access (.accdb): lectura, esquema, linked tables,
  compare, escritura controlada, seeds y teardown. Usar cuando el trabajo requiere
  inspeccionar o modificar DATOS del backend Access, no cuando el objetivo principal es
  importar/exportar VBA o UI del frontend.
---

# access-query

## Propósito

Esta skill sirve para trabajar con **datos y esquema** de backends Access (`.accdb`).

Usala cuando necesites:
- consultar datos con SQL
- listar tablas o tablas linked
- inspeccionar el esquema de una tabla
- comparar dos backends
- sembrar fixtures de test
- hacer teardown / cleanup
- ejecutar `INSERT` / `UPDATE` / `DELETE` / DDL con guardas

## No usar esta skill para

No usar `access-query` para:
- exportar/importar módulos VBA
- importar UI de formularios/reportes
- sincronizar `src/` con el binario Access
- sandbox / ERD / introspección de objetos del frontend

Para eso usar **`access-vba-sync`**.

---

## Trigger operativo para IA

Si la tarea menciona cualquiera de estos contextos, cargar esta skill:
- "consultar backend"
- "ejecutar SQL en Access"
- "listar tablas"
- "ver esquema"
- "seed / teardown / fixtures"
- "comparar datos entre backends"
- "validar que existe un registro"
- "contar filas"
- "valores distintos de un campo"

Si la tarea mezcla **código VBA/UI** y **datos backend**, usar:
1. `access-query` para inspección/seed de datos
2. `access-vba-sync` para import/export del frontend

---

## Regla de decisión rápida

| Quiero... | Usar |
|---|---|
| Ver tablas locales | `-ListTables` |
| Ver tablas linked (tabla local, source table, backend destino y connect) | `-LinkedTables` |
| Ver columnas de una tabla | `-GetSchema -Table TbX` |
| Hacer un SELECT libre | `-SQL "SELECT ..."` |
| Contar filas | `-Count -Table TbX` |
| Ver valores únicos | `-Distinct -Table TbX -Field Campo` |
| Comparar dos backends | `-Compare -CompareSQL "SELECT ..." -Backend A -CompareBackend B` |
| Ejecutar escritura inline | `-Exec "INSERT/UPDATE/DELETE ..."` |
| Ejecutar un `.sql` | `-Script ".\\fixtures\\seed.sql"` |
| Seed seguro | `-Seed -Script ".\\fixtures\\seed.sql" -AllowTable "TbX"` |
| Teardown seguro | `-Teardown -Script ".\\fixtures\\cleanup.sql" -AllowTable "TbX"` |
| Automatizar desde otra IA/script | añadir `-Json` |
| Validar sin tocar datos | añadir `-DryRun` |

---

## Reglas duras

### 1) Primero leer, después escribir
Antes de tocar datos, normalmente hacer esta secuencia:
1. `-ListTables` o `-GetSchema`
2. `-SQL` / `-Count` / `-Distinct` para entender los datos
3. recién después `-Exec`, `-Seed` o `-Teardown`

### 2) No escribir a ciegas
Si vas a escribir, preferir siempre:
- `-DryRun`
- `-AllowTable`
- `-StrictWrite` cuando corresponda

### 3) Seed/teardown exigen allow-list
`-Seed` y `-Teardown` deben llevar `-AllowTable`, salvo que el usuario pida explícitamente `-Force`.

### 4) No confundir frontend con backend
`access-query` toca **datos**. No importa formularios, no sincroniza VBA, no arregla `LoadFromText`.

---

## Rutas portables: regla oficial

La skill ahora resuelve rutas así:

1. **ruta absoluta** → se usa tal cual
2. **variables de entorno** como `%USERPROFILE%` → se expanden
3. **`~`** → se expande al perfil del usuario
4. **ruta relativa** → se intenta resolver en este orden:
   - relativa al **directorio actual** (`CWD`, normalmente la raíz del repo)
   - si no existe, relativa a la **carpeta de la skill**

### Consecuencia práctica
- En casa y en oficina, preferir `%USERPROFILE%` en `backends.json`
- Para scripts `.sql` y `-BackendPath`, preferir rutas relativas al repo cuando sea posible

### Ejemplos válidos

```powershell
# Ruta directa relativa al repo actual
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\condor_datos.accdb" -ListTables

# Ruta directa con variable de entorno
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath "%USERPROFILE%\Telefonica\Proyecto\backend.accdb" -ListTables

# Script SQL relativo al repo actual
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -Seed -Script ".\fixtures\seed_hash.sql" -AllowTable "TbSolicitudes"
```

---

## Uso canónico desde PowerShell / subproceso

### Regla crítica
Usar siempre **`-File`**, nunca `-Command`.

```powershell
powershell -ExecutionPolicy Bypass -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -SQL "SELECT * FROM TbSolicitudes"
```

No usar:

```powershell
powershell -ExecutionPolicy Bypass -Command "& '...query-backend.ps1' -SQL \"SELECT ...\""
```

---

## Comandos canónicos

### Lectura segura

```powershell
# Tablas
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -ListTables

# Linked tables
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -LinkedTables

# Linked tables en JSON detallado
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -LinkedTables -Json

# Esquema
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -GetSchema -Table "TbSolicitudes"

# SELECT
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -SQL "SELECT * FROM TbSolicitudes WHERE Estado = 'Borrador'"

# SELECT sin límite
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -SQL "SELECT * FROM TbSolicitudes" -Top 0

# Count
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -Count -Table "TbSolicitudes"

# Distinct
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -Distinct -Table "TbSolicitudes" -Field "Estado"
```

### Escritura segura

```powershell
# Dry-run antes de escribir
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -Exec "UPDATE TbSolicitudes SET Estado='Validado' WHERE ID=99901" -AllowTable "TbSolicitudes" -DryRun

# Escritura real
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -Exec "UPDATE TbSolicitudes SET Estado='Validado' WHERE ID=99901" -AllowTable "TbSolicitudes"

# Seed desde script
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -Seed -Script ".\fixtures\seed_hash.sql" -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE"

# Teardown desde script
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -Teardown -Script ".\fixtures\cleanup_hash.sql" -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE"
```

### JSON para automatización

```powershell
powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -GetSchema -Table "TbSolicitudes" -Json

powershell -File "%USERPROFILE%\.config\opencode\skills\access-query\query-backend.ps1" -BackendPath ".\backend\datos.accdb" -SQL "SELECT ID, Referencia FROM TbSolicitudes" -Top 0 -Json
```

---

## Seguridad y guardas

La skill bloquea automáticamente escrituras sobre:
1. tablas linked
2. tablas de la deny-list (`backends.json > deny_tables`)
3. tablas fuera del `-AllowTable` cuando se provee allow-list

### Recomendación de agente
- Para datos reales: usar `-DryRun` primero
- Para fixtures: usar `-Seed` / `-Teardown` con `-AllowTable`
- Para operaciones delicadas: sumar `-StrictWrite`

---

## Passwords

Orden de prioridad:
1. `-Password`
2. `ACCESS_QUERY_PW_<BACKEND>`
3. `ACCESS_QUERY_PASSWORD`
4. `.secrets.json`
5. `backends.json > password` (deprecado)
6. sin password

### Regla para IA
No hardcodear passwords en docs ni scripts temporales salvo que el usuario lo pida explícitamente.

---

## `backends.json` recomendado

### Portable entre casa y oficina

```json
{
  "default": "backend_principal",
  "deny_tables": ["MSys*", "Expediente*"],
  "backends": {
    "backend_principal": {
      "path": "%USERPROFILE%\\Telefonica\\Proyecto\\Backend.accdb",
      "description": "Backend principal"
    },
    "backend_dev": {
      "path": ".\\backend\\BackendDev.accdb",
      "description": "Backend local relativo al repo"
    }
  }
}
```

### Regla
- Preferir `%USERPROFILE%` si la ruta depende del perfil del usuario
- Preferir rutas relativas si el backend vive dentro o cerca del repo
- No guardar passwords reales en `backends.json`

---

## Qué hacer si falla

### "Ruta no encontrada"
La skill ahora muestra las rutas probadas. Revisar:
- si la ruta era relativa al repo correcto
- si convenía usar `%USERPROFILE%`
- si el archivo existe de verdad

### Error OLEDB de password
Probar en este orden:
1. `-Password`
2. env var `ACCESS_QUERY_PASSWORD`
3. `.secrets.json`

### SQL interpreta un literal como parámetro
Si Access dice algo como "No se han especificado valores para algunos parámetros requeridos", probar:
- `-Distinct` sobre ese campo para verificar valores reales
- usar comillas dobles en el literal Access SQL si hace falta

---

## Regla final para agentes

- Si necesitás **datos** o **backend** → usar `access-query`
- Si necesitás **código/UI/binario Access** → usar `access-vba-sync`
- Si la tarea mezcla ambas cosas, primero inspeccioná/sembrá con `access-query` y después cerrá el workflow de frontend con `access-vba-sync`
