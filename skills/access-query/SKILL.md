---
name: access-query
description: >
  Ejecuta SQL (lectura Y escritura) contra backends Access (.accdb) de proyectos VBA.
  Usar cuando necesites: ejecutar SELECT/INSERT/UPDATE/DELETE libre, obtener esquema
  de una tabla, listar tablas (locales o linked), contar registros, explorar valores únicos,
  sembrar fixtures de test con guardas de seguridad, ejecutar scripts .sql, o hacer cleanup.
  Incluye: deny-list con wildcards, allow-list, dry-run, -Json para automatización,
  fixture log acumulativo, y resolución de passwords sin hardcoding.
---

# ACCESS-QUERY v3 — Referencia rápida para agentes IA

> **Lee esta sección primero. Los detalles completos están más abajo.**

## ▶ DECIDE QUÉ COMANDO USAR

| Quiero... | Comando |
|-----------|---------|
| Ver qué tablas hay | `-ListTables` |
| Ver campos de una tabla | `-GetSchema -Table TbX` |
| Consultar datos | `-SQL "SELECT ..."` |
| Contar filas | `-Count -Table TbX` |
| Ver valores únicos de un campo | `-Distinct -Table TbX -Field Campo` |
| Insertar / actualizar / borrar | `-Exec "INSERT/UPDATE/DELETE ..."` |
| Ejecutar un fichero .sql | `-Script "ruta\seed.sql"` |
| Poblar datos de test (seed) | `-Seed -Script "seed.sql" -AllowTable "TbX"` |
| Limpiar datos de test (teardown) | `-Teardown -Script "clean.sql" -AllowTable "TbX"` |
| Ver tablas externas (linked) | `-LinkedTables` |
| Validar sin ejecutar | Añadir `-DryRun` a cualquier comando de escritura |
| Obtener JSON (para scripts) | Añadir `-Json` |

---

## ▶ COMANDOS DE LECTURA (sin riesgo)

```powershell
# Listar todas las tablas locales
.\query-backend.ps1 -ListTables

# Ver campos, tipos y nullabilidad de una tabla
.\query-backend.ps1 -GetSchema -Table TbSolicitudes

# SELECT libre (por defecto muestra hasta 20 filas)
.\query-backend.ps1 -SQL "SELECT * FROM TbSolicitudes WHERE Estado = 'Borrador'"

# SELECT sin límite de filas
.\query-backend.ps1 -SQL "SELECT * FROM TbSolicitudes" -Top 0

# SELECT con más filas
.\query-backend.ps1 -SQL "SELECT * FROM TbSolicitudes" -Top 100

# Contar registros
.\query-backend.ps1 -Count -Table TbSolicitudes

# Valores únicos de un campo
.\query-backend.ps1 -Distinct -Table TbSolicitudes -Field Estado

# Tablas linked (externas, solo lectura)
.\query-backend.ps1 -LinkedTables
```

### Lectura con salida JSON (para automatización)

```powershell
# GetSchema → JSON
.\query-backend.ps1 -GetSchema -Table TbSolicitudes -Json | ConvertFrom-Json

# SQL → JSON con todas las filas
.\query-backend.ps1 -SQL "SELECT ID, Referencia FROM TbSolicitudes" -Top 0 -Json | ConvertFrom-Json

# ListTables → JSON
.\query-backend.ps1 -ListTables -Json | ConvertFrom-Json
```

---

## ▶ COMANDOS DE ESCRITURA

### Exec — SQL inline

```powershell
# INSERT simple
.\query-backend.ps1 -Exec "INSERT INTO TbSolicitudes (ID, Referencia, Estado) VALUES (99901, 'TEST_01', 'Borrador')"

# UPDATE
.\query-backend.ps1 -Exec "UPDATE TbSolicitudes SET Estado='Validado' WHERE ID=99901"

# DELETE
.\query-backend.ps1 -Exec "DELETE FROM TbSolicitudes WHERE ID=99901"

# Multi-sentencia (parser respeta ; dentro de strings)
.\query-backend.ps1 -Exec "INSERT INTO TbSol (ID, Obs) VALUES (99901, 'tiene;punto_y_coma'); INSERT INTO TbSol (ID) VALUES (99902)"

# Dry-run: validar guardas sin ejecutar
.\query-backend.ps1 -Exec "INSERT INTO TbSolicitudes (ID) VALUES (99901)" -DryRun
```

### Script — desde fichero .sql

```powershell
.\query-backend.ps1 -Script ".\fixtures\seed_datos.sql"
.\query-backend.ps1 -Script ".\fixtures\seed_datos.sql" -DryRun
```

### Seed / Teardown — fixtures con tracking

```powershell
# Seed desde fichero (requiere -AllowTable)
.\query-backend.ps1 -Seed -Script ".\fixtures\seed_hash.sql" -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE"

# Seed inline
.\query-backend.ps1 -Seed -Exec "INSERT INTO TbSolicitudes (ID, Ref) VALUES (99101, 'TEST')" -AllowTable "TbSolicitudes"

# Teardown
.\query-backend.ps1 -Teardown -Script ".\fixtures\cleanup_hash.sql" -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE"

# Teardown inline
.\query-backend.ps1 -Teardown -Exec "DELETE FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110" -AllowTable "TbSolicitudes"
```

---

## ▶ INVOCAR DESDE cmd.exe / SUBPROCESO

> ⚠️ **Regla crítica: usa siempre `-File`, nunca `-Command`.**

Con `-Command` las comillas anidadas se rompen al pasar por cmd.exe y el binding de parámetros falla. Con `-File` los argumentos se pasan directamente sin reinterpretación.

```bat
:: ✅ CORRECTO — usar -File
powershell -ExecutionPolicy Bypass -File "C:\ruta\access-query\query-backend.ps1" -Password "dpddpd" -BackendPath "C:\ruta\archivo.accdb" -SQL "SELECT Campo FROM Tabla WHERE Tipo = 'VALOR'"

:: ❌ INCORRECTO — nunca usar -Command con este script
powershell -ExecutionPolicy Bypass -Command "& 'C:\ruta\query-backend.ps1' -SQL \"SELECT ...\""
```

Con `-File` las comillas simples dentro del SQL (`WHERE x = 'valor'`) funcionan sin escape adicional.

### Error "No se han especificado valores para algunos de los parámetros requeridos"

Este error de Access OleDb **no es un problema de comillas** — significa que Access está interpretando un literal de texto como un parámetro de consulta. Ocurre cuando el valor en el `WHERE` coincide con el nombre de un campo u objeto de la BD.

Solución: usar comillas dobles como delimitador de string en lugar de simples (sintaxis alternativa válida en Access SQL):

```bat
powershell -ExecutionPolicy Bypass -File "...\query-backend.ps1" -Password "dpddpd" -BackendPath "ruta.accdb" -SQL "SELECT Campo FROM Tabla WHERE Tipo = \"VALOR\""
```

Si persiste, verificar primero que el valor existe con `-Distinct`:

```bat
powershell -ExecutionPolicy Bypass -File "...\query-backend.ps1" -Password "dpddpd" -BackendPath "ruta.accdb" -Distinct -Table "Tabla" -Field "Tipo"
```

---

## ▶ FLUJO COMPLETO DE FIXTURE (referencia rápida)

```powershell
# La password ya está configurada en .secrets.json — no hace falta setup previo

# 1. Ver qué tablas hay
.\query-backend.ps1 -ListTables

# 2. Inspeccionar tabla objetivo
.\query-backend.ps1 -GetSchema -Table TbSolicitudes

# 3. Dry-run del seed
.\query-backend.ps1 -Seed -Script ".\fixtures\seed_hash.sql" -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE" -DryRun

# 4. Ejecutar seed
.\query-backend.ps1 -Seed -Script ".\fixtures\seed_hash.sql" -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE"

# 5. Verificar
.\query-backend.ps1 -SQL "SELECT ID, Referencia FROM TbSolicitudes WHERE ID BETWEEN 99101 AND 99110"

# 6. Teardown
.\query-backend.ps1 -Teardown -Script ".\fixtures\cleanup_hash.sql" -AllowTable "TbSolicitudes" -FixtureTag "HASH_EDGE"

# 7. Confirmar limpieza
.\query-backend.ps1 -Count -Table TbSolicitudes
```

---

## ▶ SELECCIONAR BACKEND

```powershell
# Backend por defecto (definido en backends.json > "default")
.\query-backend.ps1 -ListTables

# Backend específico por alias
.\query-backend.ps1 -ListTables -Backend backend_cache

# Ruta directa (sin backends.json)
.\query-backend.ps1 -ListTables -BackendPath "C:\ruta\al\archivo.accdb"
```

---

## ▶ PASSWORDS

**La contraseña estándar de todos los backends de este proyecto es `dpddpd`.**

El `.secrets.json` ya está configurado con ella. No necesitas hacer nada — simplemente ejecuta los comandos y funcionará.

Si por alguna razón necesitas sobreescribirla puntualmente:

```powershell
# Override puntual (solo para ese comando)
.\query-backend.ps1 -ListTables -Password "otra_password"

# Override por sesión completa
$env:ACCESS_QUERY_PASSWORD = "otra_password"
```

Cadena de prioridad completa (de mayor a menor):
`-Password` > env `ACCESS_QUERY_PW_<BACKEND>` > env `ACCESS_QUERY_PASSWORD` > `.secrets.json` > `backends.json`

---

## ▶ SEGURIDAD — Lo que bloquea escrituras automáticamente

1. **Tablas linked** — jamás se puede escribir en ellas (protección automática)
2. **Deny-list** — tablas en `backends.json > deny_tables` (soporta wildcards `Expediente*`)
3. **Allow-list** — si se pasa `-AllowTable`, solo esas tablas aceptan escritura
4. **`-Seed` / `-Teardown`** — **exigen** `-AllowTable` (salvo `-Force`)

```powershell
# -StrictWrite: exige -AllowTable en CUALQUIER modo de escritura
.\query-backend.ps1 -Exec "UPDATE TbSol SET X=1" -StrictWrite -AllowTable "TbSol"

# -Force: saltar requisito de -AllowTable (no usar en producción)
.\query-backend.ps1 -Seed -Exec "INSERT INTO TbSol (ID) VALUES (1)" -Force
```

Si una sentencia es bloqueada, **toda la ejecución se aborta** (fail-fast).

---

## ▶ ESTRUCTURA DE ARCHIVOS

```
.agents/skills/access-query/
├── SKILL.md                 ← este archivo (NO modificar)
├── query-backend.ps1        ← script único para todos los modos (NO modificar)
├── backends.json            ← mapa de backends + deny-list (SIN passwords)
├── .secrets.json.template   ← plantilla para secrets
├── .secrets.json            ← passwords reales (NO versionar, en .gitignore)
├── .fixture-log.json        ← log acumulativo de fixtures (NO versionar)
└── fixtures/                ← scripts .sql de seed/teardown
```

> ⛔ **PROHIBIDO para la IA:**
> - Crear o modificar cualquier archivo dentro de `.agents/skills/access-query/`
> - Generar scripts `.ps1` auxiliares o wrappers dentro de esta carpeta
> - No necesitas ningún script adicional: `query-backend.ps1` cubre todos los casos
>
> ✅ **Si necesitas un script temporal** (seed ad-hoc, consulta compuesta, etc.):
> créalo en la **raíz del repositorio** en el que estés trabajando, no aquí.

Añadir a `.gitignore`:
```
.secrets.json
.fixture-log.json
```

---

## ▶ CONFIGURAR backends.json

```json
{
  "default": "backend_principal",
  "deny_tables": [
    "Expedientes", "Expediente*",
    "NoConformidades", "Lanzadera*", "MSys*"
  ],
  "backends": {
    "backend_principal": {
      "path": "C:\\ruta\\al\\Backend_Datos.accdb",
      "description": "Backend con datos reales"
    },
    "backend_dev": {
      "path": "C:\\ruta\\al\\Backend_Dev.accdb",
      "description": "Backend de desarrollo"
    }
  }
}
```

El campo `password` en backends.json está **deprecado**. Usar `.secrets.json` o env vars.

---

## ▶ REFERENCIA DE PARÁMETROS

### Lectura

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `-SQL` | string | SELECT libre |
| `-Top` | int | Límite de filas (default: 20; **0 = sin límite**) |
| `-Table` | string | Para -GetSchema, -Count, -Distinct |
| `-Field` | string | Para -Distinct |
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
| `-AllowTable` | string | CSV de tablas permitidas (p.ej. `"TbSol,TbDocs"`) |
| `-DenyTable` | string | CSV de tablas bloqueadas adicionales |
| `-StrictWrite` | switch | Requiere -AllowTable en TODO modo de escritura |
| `-Force` | switch | Salta requisito de AllowTable en Seed/Teardown |

### Conexión y salida

| Parámetro | Tipo | Descripción |
|-----------|------|-------------|
| `-Backend` | string | Alias de backend (de backends.json) |
| `-BackendPath` | string | Ruta directa (ignora backends.json) |
| `-Password` | string | Override de password |
| `-Json` | switch | Emite JSON estructurado a stdout (lectura Y escritura) |

---

## ▶ SALIDA JSON — Estructura

### Lectura (SQL / GetSchema / ListTables con -Json)

```json
{
  "mode": "SQL",
  "backend": "backend_principal",
  "sql": "SELECT ID, Referencia FROM TbSolicitudes WHERE ID > 99100",
  "rowCount": 6,
  "columns": ["ID", "Referencia", "Estado"],
  "rows": [
    { "ID": 99101, "Referencia": "EDGE_HASH_NULL", "Estado": "Borrador" },
    { "ID": 99102, "Referencia": "EDGE_HASH_EMPTY", "Estado": "Borrador" }
  ]
}
```

### GetSchema con -Json

```json
{
  "mode": "GetSchema",
  "table": "TbSolicitudes",
  "backend": "backend_principal",
  "columns": [
    { "name": "ID", "type": "Integer", "nullable": false },
    { "name": "Referencia", "type": "String(100)", "nullable": true }
  ]
}
```

### ListTables con -Json

```json
{
  "mode": "ListTables",
  "backend": "backend_principal",
  "tables": ["TbSolicitudes", "TbDocumentos", "TbActividades"]
}
```

### Escritura (Exec / Seed / Teardown con -Json)

```json
{
  "mode": "SEED [HASH_EDGE]",
  "backend": "backend_principal",
  "dryRun": false,
  "fixtureTag": "HASH_EDGE",
  "timestamp": "2025-12-01T10:30:00.000+01:00",
  "statements": [
    { "index": 1, "sql": "INSERT INTO ...", "type": "WRITE", "status": "OK", "targets": ["TbSolicitudes"], "affected": 1 },
    { "index": 2, "sql": "SELECT ...", "type": "READ", "status": "OK", "affected": 6 }
  ],
  "aborted": false,
  "totalAffected": 5,
  "tablesWritten": ["TbSolicitudes"]
}
```

---

## ▶ SQL DE ACCESS — Recordatorio rápido

```sql
-- Wildcard en LIKE: * en lugar de %
SELECT * FROM TbSolicitudes WHERE Nombre LIKE "Mar*"

-- Fechas con #
SELECT * FROM TbActividades WHERE Fecha >= #01/15/2024#

-- DateAdd
SELECT * FROM TbActividades WHERE Fecha >= DateAdd("m", -3, DATE())

-- Escapar comilla simple
INSERT INTO TbSolicitudes (Nombre) VALUES ('O''Brien')

-- Nombres con espacios: corchetes
SELECT * FROM [Mi Tabla Con Espacios]

-- No hay TRUNCATE: usar DELETE
DELETE FROM TbSolicitudes

-- Count y group
SELECT Oficina, COUNT(*) FROM TbSolicitudes GROUP BY Oficina
```

---

## ▶ CONVENCIONES PARA FIXTURES

| Prefijo de Referencia | Uso |
|-----------|-----|
| `TEST_` | Fixtures genéricos |
| `EDGE_` | Edge cases |
| `ZZZ_`  | Tablas temporales |

| Rango ID | Uso |
|----------|-----|
| 99001–99099 | Integración |
| 99101–99199 | Hash/validación |
| 99801–99899 | Exportación |
| 99901–99999 | Ad-hoc |

---

## ▶ LIMITACIONES CONOCIDAS

1. **No hay transacciones reales en Access OleDb.** Si la sentencia 3 de 5 falla en runtime, las 2 primeras ya se ejecutaron. El fail-fast de guardas es preventivo (valida antes), pero errores SQL en runtime no se deshacen.

2. **El parser de tablas usa regex.** Cubre INSERT/UPDATE/DELETE/DROP/CREATE/ALTER. Subqueries complejas con JOINs anidados podrían no detectar la tabla correcta.

3. **No soporta bloques `/* */`** en comentarios SQL. Usar `--` en su lugar.

4. **`-Json` en lectura** emite a stdout. `Write-Host` va a stderr/consola y no interfiere. Usar `| ConvertFrom-Json` para parsear.

5. **`.fixture-log.json` es append-only.** Limpiar manualmente si crece mucho.
