---
name: access-query
description: >
  Ejecuta SQL (lectura Y escritura) contra backends Access (.accdb) de proyectos VBA.
  Usar cuando necesites: ejecutar SQL libre (SELECT, INSERT, UPDATE, DELETE), obtener el esquema
  de una tabla, listar tablas (locales o linked), contar registros, explorar valores únicos,
  comparar resultados entre backends, sembrar fixtures de test con guardas de seguridad,
  ejecutar scripts .sql, o hacer cleanup de datos de sandbox. Incluye: bloqueo de tablas linked,
  deny-list con wildcards, allow-list, dry-run, -StrictWrite, -Json para automatización,
  fixture log acumulativo, y resolución de passwords sin hardcoding.
---

# ACCESS-QUERY v2 — Consultas y Escritura Segura a Backends Access

## ⚠️ Antes de usar: configurar passwords

**No dejes passwords hardcodeadas en `backends.json`.**

Cadena de prioridad para resolución de password:
1. `-Password "pw"` — override puntual desde CLI
2. `ACCESS_QUERY_PW_<BACKEND>` — env var por backend (ej: `ACCESS_QUERY_PW_backend_principal`)
3. `ACCESS_QUERY_PASSWORD` — env var global
4. `.secrets.json` — fichero local en el directorio de la skill (NO versionar)
5. `backends.json > password` — backward compat, **DEPRECADO**

**Setup recomendado:**
```powershell
# Opcion A: env var por sesion
$env:ACCESS_QUERY_PASSWORD = "tu_password"

# Opcion B: .secrets.json (copiar desde .secrets.json.template)
# { "backend_principal": "tu_password", "backend_cache": "otra_password" }

# Opcion C: env var persistente (solo tu perfil)
[Environment]::SetEnvironmentVariable('ACCESS_QUERY_PASSWORD', 'tu_password', 'User')
```

Agregar a `.gitignore`:
```
.secrets.json
.fixture-log.json
```

---

## Prerequisitos

- `Microsoft.ACE.OLEDB.12.0` instalado
- Ruta del `.accdb` accesible desde la sesión actual
- `backends.json` con rutas y deny-list configuradas
- Password accesible por alguna de las vías anteriores

---

## Archivos de la skill

```
.agents/skills/access-query/
├── SKILL.md                 ← este archivo
├── query-backend.ps1        ← script único para todos los modos
├── backends.json            ← mapa de backends + deny-list (SIN passwords)
├── .secrets.json.template   ← plantilla para el fichero de secretos
├── .secrets.json            ← passwords reales (NO versionar)
├── .fixture-log.json        ← log acumulativo de fixtures (NO versionar)
└── fixtures/                ← scripts .sql de seed/teardown
```

---

## Configurar backends.json

```json
{
  "default": "backend_principal",
  "deny_tables": [
    "Expedientes", "Expediente*",
    "NoConformidades", "Lanzadera*", "MSys*"
  ],
  "backends": {
    "backend_principal": {
      "path": "C:\\ruta\\al\\Backend.accdb",
      "description": "Backend con datos reales"
    }
  }
}
```

El campo `password` ya no es necesario en `backends.json`. Usa `.secrets.json` o env vars.

---

## Referencia de parámetros

### Lectura

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `-SQL` | string | SELECT libre |
| `-Table` | string | Para -GetSchema, -Count, -Distinct |
| `-Field` | string | Para -Distinct |
| `-Top` | int | Límite de filas en -SQL (default: 20) |
| `-Count` | switch | Contar registros |
| `-Distinct` | switch | Valores únicos |
| `-ListTables` | switch | Tablas locales |
| `-LinkedTables` | switch | Tablas linked |
| `-GetSchema` | switch | Esquema de campos |
| `-Compare` | switch | Comparar dos backends |
| `-CompareSQL` | string | SQL para comparación |
| `-CompareBackend` | string | Segundo backend |

### Escritura

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `-Exec` | string | SQL inline (multi-sentencia con `;`) |
| `-Script` | string | Ruta a fichero `.sql` |
| `-Seed` | switch | Modo seed: etiqueta + log de fixture |
| `-Teardown` | switch | Modo teardown: etiqueta + log |
| `-FixtureTag` | string | Tag identificador (default: `FX_yyyyMMdd_HHmmss`) |
| `-CreateTable` | switch | DDL creación (usa -Exec) |
| `-DropTable` | switch | DDL drop (usa -Table) |

### Seguridad

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `-DryRun` | switch | Validar guardas sin ejecutar |
| `-AllowTable` | string | CSV de tablas permitidas |
| `-DenyTable` | string | CSV de tablas bloqueadas (suma a backends.json) |
| `-StrictWrite` | switch | Requiere -AllowTable en TODO modo de escritura |
| `-Force` | switch | Salta requisito de AllowTable en Seed/Teardown |

### Salida

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `-Json` | switch | Emite JSON estructurado a stdout |

### Conexión

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `-Backend` | string | Alias de backend |
| `-BackendPath` | string | Ruta directa (ignora backends.json) |
| `-Password` | string | Override de password |

---

## Modelo de seguridad

### Tres capas para escritura

1. **Linked detection**: `GetSchema('Tables')` detecta tablas con `TABLE_TYPE = 'LINK'`. Escritura bloqueada.
2. **Deny-list**: `backends.json > deny_tables` + `-DenyTable`. Soporta wildcards (`Expediente*`).
3. **Allow-list**: Si `-AllowTable` está presente, solo esas tablas aceptan escritura.

### -Seed / -Teardown requieren -AllowTable

Para evitar que un seed toque una tabla local equivocada por accidente, estos modos **exigen** `-AllowTable`. Usar `-Force` para saltarse esta restricción (no recomendado en producción).

### -StrictWrite

Extiende el requisito de `-AllowTable` a **todos** los modos de escritura (-Exec, -Script, -DDL).

### Fail-fast

Si cualquier sentencia es bloqueada, **toda la ejecución se aborta**. Las sentencias posteriores no se ejecutan.

---

## Fixture management real

### Qué hacen -Seed y -Teardown

Más allá de etiquetar la salida, estos modos:

1. **Exigen `-AllowTable`** — primera guarda extra
2. **Generan `-FixtureTag`** automático si no se pasa (`FX_yyyyMMdd_HHmmss`)
3. **Escriben en `.fixture-log.json`** — log acumulativo con tag, timestamp, backend, tablas, filas afectadas
4. **Emiten JSON estructurado** con `-Json` — consumible por scripts de test

### .fixture-log.json

```json
[
  {
    "fixtureTag": "HASH_EDGE_CASES",
    "mode": "SEED [HASH_EDGE_CASES]",
    "timestamp": "2025-12-01T10:30:00.000+01:00",
    "backend": "backend_principal",
    "allowList": ["TbSolicitudes"],
    "dryRun": false,
    "aborted": false,
    "tables": ["TbSolicitudes"],
    "affected": 5,
    "stmtCount": 6
  },
  {
    "fixtureTag": "HASH_EDGE_CASES",
    "mode": "TEARDOWN [HASH_EDGE_CASES]",
    "timestamp": "2025-12-01T11:00:00.000+01:00",
    "backend": "backend_principal",
    "allowList": ["TbSolicitudes"],
    "dryRun": false,
    "aborted": false,
    "tables": ["TbSolicitudes"],
    "affected": 5,
    "stmtCount": 2
  }
]
```

**Nota:** Ambos registros comparten `fixtureTag: "HASH_EDGE_CASES"`, lo que permite trazabilidad completa seed→teardown.

### Salida -Json

```json
{
  "mode": "SEED [HASH_EDGE_CASES]",
  "backend": "backend_principal",
  "dryRun": false,
  "fixtureTag": "HASH_EDGE_CASES",
  "timestamp": "2025-12-01T10:30:00.000+01:00",
  "statements": [
    { "index": 1, "sql": "INSERT INTO ...", "type": "WRITE", "status": "OK", "targets": ["TbSolicitudes"], "affected": 1 },
    { "index": 2, "sql": "SELECT ...", "type": "READ", "status": "OK", "affected": 1 }
  ],
  "aborted": false,
  "totalAffected": 1,
  "tablesWritten": ["TbSolicitudes"]
}
```

---

## Modos de uso

### Lectura (sin cambios)

```powershell
.\query-backend.ps1 -SQL "SELECT TOP 10 * FROM TbSolicitudes"
.\query-backend.ps1 -GetSchema -Table "TbSolicitudes"
.\query-backend.ps1 -ListTables
.\query-backend.ps1 -LinkedTables
```

### Exec — escritura inline

```powershell
.\query-backend.ps1 -Exec "INSERT INTO TbSolicitudes (ID, Referencia) VALUES (99901, 'TEST_01')"

# Multi-sentencia (el parser respeta ; dentro de strings)
.\query-backend.ps1 -Exec "INSERT INTO TbSol (ID, Obs) VALUES (99901, 'tiene;punto_y_coma'); INSERT INTO TbSol (ID, Obs) VALUES (99902, 'normal')"

# Con StrictWrite
.\query-backend.ps1 -Exec "UPDATE TbSol SET Estado='OK' WHERE ID=99901" -StrictWrite -AllowTable "TbSol"
```

### Script — desde fichero .sql

```powershell
.\query-backend.ps1 -Script ".\fixtures\seed_hash.sql"
.\query-backend.ps1 -Script ".\fixtures\seed_hash.sql" -DryRun
```

### Seed — fixture con tracking

```powershell
# Requiere -AllowTable
.\query-backend.ps1 -Seed -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE" `
    -Exec "INSERT INTO TbSolicitudes (ID, Ref, Hash) VALUES (99101, 'EDGE_NULL', NULL)"

# Desde fichero
.\query-backend.ps1 -Seed -AllowTable "TbSolicitudes" -Script ".\fixtures\seed_hash.sql" -FixtureTag "HASH_V2"

# Con JSON para automatización
.\query-backend.ps1 -Seed -AllowTable "TbSolicitudes" -Script ".\fixtures\seed_hash.sql" -Json | ConvertFrom-Json
```

### Teardown — cleanup con tracking

```powershell
.\query-backend.ps1 -Teardown -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE" `
    -Exec "DELETE FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110"
```

### DryRun + Json — validar plan antes de ejecutar

```powershell
$plan = .\query-backend.ps1 -Seed -AllowTable "TbSolicitudes" -Script ".\fixtures\seed_hash.sql" -DryRun -Json | ConvertFrom-Json
$plan.statements | Format-Table index, type, status, targets
```

---

## Ejemplo real completo: seed de hash en CONDOR

```powershell
# 0. Setup de password (una vez por sesion)
$env:ACCESS_QUERY_PASSWORD = "tu_password"

# 1. Dry-run: ver el plan sin ejecutar
.\query-backend.ps1 -Seed -AllowTable "TbSolicitudes" `
    -Script ".\fixtures\seed_hash_validation.sql" `
    -FixtureTag "HASH_EDGE" -DryRun

# 2. Seed real
.\query-backend.ps1 -Seed -AllowTable "TbSolicitudes" `
    -Script ".\fixtures\seed_hash_validation.sql" `
    -FixtureTag "HASH_EDGE"

# 3. Verificar
.\query-backend.ps1 -SQL "SELECT ID, Referencia, HashValidacion FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110"

# 4. Ejecutar tests manuales o automatizados...

# 5. Teardown
.\query-backend.ps1 -Teardown -AllowTable "TbSolicitudes" `
    -Script ".\fixtures\cleanup_hash_validation.sql" `
    -FixtureTag "HASH_EDGE"

# 6. Confirmar limpieza
.\query-backend.ps1 -SQL "SELECT COUNT(*) FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110"
```

### Automatización con JSON

```powershell
$env:ACCESS_QUERY_PASSWORD = "tu_password"

# Seed + captura de resultado
$seedResult = .\query-backend.ps1 -Seed -AllowTable "TbSolicitudes" `
    -Script ".\fixtures\seed_hash_validation.sql" `
    -FixtureTag "HASH_AUTO" -Json | ConvertFrom-Json

if ($seedResult.aborted) {
    Write-Error "Seed fallido: $($seedResult.statements | Where-Object {$_.status -eq 'BLOCKED'} | ForEach-Object {$_.blocked})"
    exit 1
}

Write-Host "Seed OK: $($seedResult.totalAffected) filas en $($seedResult.tablesWritten -join ', ')"

# ... ejecutar tests ...

# Teardown
$cleanResult = .\query-backend.ps1 -Teardown -AllowTable "TbSolicitudes" `
    -Exec "DELETE FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110" `
    -FixtureTag "HASH_AUTO" -Json | ConvertFrom-Json

Write-Host "Cleanup: $($cleanResult.totalAffected) filas eliminadas"
```

---

## Notas de Access SQL

- Wildcard en LIKE: `*` en lugar de `%` → `WHERE Nombre LIKE "Mar*"`
- Fechas: `DATE()`, `#2024-01-15#`, `DateAdd("m", -3, DATE())`
- Corchetes para nombres especiales: `[Mi Tabla]`, `[Mi Campo]`
- No hay `TRUNCATE TABLE`: usar `DELETE FROM [tabla]`
- `''` (dos comillas simples) para escapar `'` dentro de un string

---

## Convenciones para fixtures

| Prefijo | Uso |
|---------|-----|
| `TEST_` | Fixtures genéricos |
| `EDGE_` | Edge cases |
| `ZZZ_`  | Tablas temporales |
| `FX_`   | Tag generado automáticamente |

| Rango ID | Uso |
|----------|-----|
| 99001–99099 | Integración |
| 99101–99199 | Hash/validación |
| 99801–99899 | Exportación |
| 99901–99999 | Ad-hoc |

---

## Limitaciones conocidas

1. **Access OleDb no tiene transacciones reales.** Si la sentencia 3 de 5 falla, las 2 primeras ya se ejecutaron. El fail-fast de guardas es preventivo (valida antes), pero errores SQL en runtime no se deshacen.

2. **El parser de tablas usa regex.** Cubre los patrones estándar (INSERT INTO, UPDATE, DELETE FROM, DROP/CREATE TABLE). SQL con subqueries anidadas complejas podría no detectar la tabla correcta.

3. **El parser de sentencias respeta `'`, `"` y `--`.** No soporta bloques `/* */` ni sentencias con heredocs. Para fixtures de test esto es más que suficiente.

4. **`-Json` emite a stdout.** `Write-Host` va a stderr/consola y no interfiere. Usar `| ConvertFrom-Json` para parsear.

5. **`.fixture-log.json` es append-only.** No se purga automáticamente. Limpiar manualmente o con script si crece mucho.

---

## Parser de SQL — Alcance y límites

### Lo que soporta bien

El parser de `query-backend.ps1` divide un bloque SQL en sentencias individuales respetando:

- **Punto y coma `;`** como separador (fuera de strings)
- **Comillas simples `'`** con escape `''` ( Access SQL standard )
- **Comillas dobles `"`** ( Access permite identificadores entre comillas )
- **Comentarios inline `--`** (hasta fin de línea, fuera de strings)

**Ejemplos de patrones bien reconocidos:**

```sql
-- Insert con punto y coma interno en string
INSERT INTO TbSolicitudes (ID, Obs) VALUES (99901, 'tiene;punto_y_coma');

-- Comentario antes de sentencia
-- Este es un comentario
INSERT INTO TbSol (ID, Nombre) VALUES (99902, 'TEST');

-- Multi-sentencia en una línea
INSERT INTO TbSol (ID) VALUES (99903); INSERT INTO TbSol (ID) VALUES (99904);

-- Update con condición
UPDATE TbSol SET Estado='OK' WHERE ID=99901;

-- Delete con comentario al lado
DELETE FROM TbSol WHERE ID=99902; -- cleanup post-test

-- String con comillas simples dobladas
INSERT INTO TbSol (ID, Nombre) VALUES (99905, 'O''Brien');

-- Identifier con espacio entre corchetes
INSERT INTO [Mi Tabla] (ID) VALUES (1);
```

### Lo que NO soporta (o puede fallar)

- **Bloques `/* ... */`** — el parser los trata como contenido normal (el `*/` no cierra el bloque)
- **Subqueries complejas** donde la tabla real está en un JOIN profundo — el parser extrae la primera tabla que encuentra, que podría no ser la de escritura
- **MERGE, UPSERT** — no reconocidos por el regex de escritura
- **Sentencias con -- dentro de strings**: `INSERT INTO T (C) VALUES ('abc -- def')` — el parser correctamente interpreta el `'--'` dentro del string como literal, pero si hay `--` fuera del string puede truncar prematuramente
- **Nombres de tabla con punto `.` en el nombre**: `[dbo].[MiTabla]` — el regex captura hasta el `.` como parte del nombre

### Ejemplos de Access SQL real

```sql
-- Wildcard * en LIKE (no %)
SELECT * FROM TbSolicitudes WHERE Nombre LIKE "Mar*"

-- Fecha con # y format MM/DD/AAAA
SELECT * FROM TbActividades WHERE Fecha >= #01/15/2024#

-- Función Date() y DateAdd
SELECT * FROM TbActividades WHERE Fecha >= DateAdd("m", -3, DATE())

-- Escapar comilla simple con ''
INSERT INTO TbSolicitudes (Nombre) VALUES ('O''Brien')

-- Corchetes para nombres especiales
SELECT * FROM [Mi Tabla Con Espacios]
INSERT INTO [Tb Externo] (ID) VALUES (1)

-- Count y grouping
SELECT Oficina, COUNT(*) FROM TbSolicitudes GROUP BY Oficina

-- Distinct
SELECT DISTINCT Estado FROM TbSolicitudes WHERE Estado IS NOT NULL
```

### Recomendación para fixtures de test

Usar siempre **comillas simples** para strings literals y preferir siempre que sea posible la forma más simple de cada statement. Los fixtures típicos de seed/teardown son INSERT/UPDATE/DELETE directos, para los cuales el parser funciona correctamente.
