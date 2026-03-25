---
name: access-query
description: >
  Ejecuta SQL y consultas de inspección contra backends Access (.accdb) de proyectos VBA.
  Usar cuando necesites: ejecutar SQL libre contra un .accdb, obtener el esquema de una tabla,
  listar tablas (locales o linked), contar registros, explorar valores únicos de un campo,
  o comparar resultados entre dos backends. Un único script cubre todos los casos.
---

# ACCESS-QUERY — Consultas a Backends Access (.accdb)

## ⚠️ Antes de usar: configurar passwords

Los backends Access de este proyecto tienen password de VBA.

**No intentes adivinar la password.** Si `backends.json` tiene password vacía o el primer
intento falla con error "No es una contraseña válida", **PREGUNTA al usuario** cuál es
la contraseña antes de reintentar.

> **Regla:** Si un `Import` o query falla porque el `.accdb` requiere password, es
> porque la password no está en `backends.json` — hay que agregarla, no buscar workarounds.

---

## Prerequisitos

Antes de ejecutar el script por primera vez, verificar que el entorno cumple:

- **`Microsoft.ACE.OLEDB.12.0` instalado** en la máquina. Es el proveedor OleDB para `.accdb`. Si no está, el script falla con `"El proveedor no está registrado en el equipo local"`. Solución: instalar el [Microsoft Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920) — usar la versión de 32 o 64 bits según el PowerShell que se use.
- **Ruta del `.accdb` accesible** desde la sesión actual. En entornos de agente (Trae, OpenCode), el proceso puede correr en un contexto diferente al del usuario interactivo. Si la ruta es de red (`\\servidor\...`), verificar que la unidad está mapeada en esa sesión.
- **`backends.json` configurado** con rutas y passwords correctas antes del primer uso (ver sección siguiente).

---

## Cuándo usar esta skill

- Ejecutar una SQL contra un backend `.accdb` antes de implementarla en VBA
- Ver los campos y tipos de una tabla (`-GetSchema`)
- Listar las tablas locales o linked de un `.accdb`
- Contar registros de una tabla
- Explorar valores distintos de un campo para preparar tests o validar datos
- Comparar el resultado de una misma SQL entre dos backends (p.ej. backend real vs. caché)

## Archivos de la skill

```
.agents/skills/access-query/
├── SKILL.md           ← este archivo
├── query-backend.ps1  ← script único para todos los modos
└── backends.json      ← mapa de backends del proyecto (editar por proyecto)
```

## Configurar backends.json

**Este fichero es específico de cada proyecto.** Edítalo antes de usar la skill.  
El campo `"default"` indica qué backend se usa cuando no se pasa `-Backend`.

```json
{
  "default": "backend_principal",
  "backends": {
    "backend_principal": {
      "path": "C:\\ruta\\al\\proyecto\\NombreBackend.accdb",
      "password": "tu_password",
      "description": "Backend principal con datos reales"
    },
    "backend_cache": {
      "path": "C:\\ruta\\al\\proyecto\\NombreFrontend.accdb",
      "password": "tu_password",
      "description": "Frontend con tablas en caché"
    }
  }
}
```

> La password de cada backend viene del `backends.json`. Usar `-Password` solo para sobreescribirla puntualmente.

---

## Referencia completa de parámetros

| Parámetro         | Tipo   | Descripción |
|-------------------|--------|-------------|
| `-SQL`            | string | SQL libre a ejecutar |
| `-Table`          | string | Nombre de tabla (para `-GetSchema`, `-Count`, `-Distinct`) |
| `-Field`          | string | Nombre de campo (para `-Distinct`) |
| `-Top`            | int    | Límite de filas en modo `-SQL` (default: 20) |
| `-Count`          | switch | Contar registros de `-Table` |
| `-Distinct`       | switch | Valores únicos de `-Field` en `-Table` |
| `-ListTables`     | switch | Listar tablas locales del backend |
| `-LinkedTables`   | switch | Listar tablas linked del backend |
| `-GetSchema`      | switch | Esquema de campos de `-Table` |
| `-Compare`        | switch | Comparar resultados de `-CompareSQL` entre dos backends |
| `-CompareSQL`     | string | SQL a comparar (usada con `-Compare`) |
| `-CompareBackend` | string | Segundo backend para la comparación |
| `-Backend`        | string | Alias del backend a usar (definido en `backends.json`). Default: el marcado como `"default"` |
| `-BackendPath`    | string | Ruta directa a un `.accdb` (ignora `backends.json`) |
| `-Password`       | string | Sobreescribe la password del backend para esta ejecución |

---

## Modos de uso

### 1. SQL libre

Ejecuta cualquier SELECT contra el backend. Muestra hasta 20 filas por defecto.

```powershell
# Backend por defecto
.\query-backend.ps1 -SQL "SELECT TOP 10 * FROM TbClientes"

# Backend específico
.\query-backend.ps1 -SQL "SELECT COUNT(*) FROM TbPedidos" -Backend backend_principal

# Limitar filas mostradas
.\query-backend.ps1 -SQL "SELECT * FROM TbClientes" -Top 5

# Ruta directa sin backends.json
.\query-backend.ps1 -SQL "SELECT * FROM TbOtro" -BackendPath "C:\ruta\otro.accdb"
```

**Salida:** cada fila como `Campo1=valor | Campo2=valor | ...` + total al final.

---

### 2. Esquema de tabla (`-GetSchema`)

Muestra campos, tipo y si admite NULL. El ancho de columna se adapta al nombre más largo.

```powershell
.\query-backend.ps1 -GetSchema -Table "TbClientes"
.\query-backend.ps1 -GetSchema -Table "TbPedidos" -Backend backend_principal
.\query-backend.ps1 -GetSchema -Table "TbCache" -Backend backend_cache
```

**Salida:**
```
=== ESQUEMA: TbClientes (backend_principal) ===

  | Campo                | Tipo           | Nullable |
  | -------------------- | -------------- | -------- |
  | ID                   | Integer        | No       |
  | NIF                  | String(20)     | Yes      |
  | Nombre               | String(100)    | Yes      |
  | F_Alta               | Date           | Yes      |
  | Activo               | Boolean        | No       |
```

---

### 3. Listar tablas

```powershell
# Tablas locales del backend por defecto
.\query-backend.ps1 -ListTables

# Tablas locales de un backend concreto
.\query-backend.ps1 -ListTables -Backend backend_cache

# Tablas linked (con su ruta de origen)
.\query-backend.ps1 -LinkedTables
.\query-backend.ps1 -LinkedTables -Backend backend_principal
```

---

### 4. Contar registros (`-Count`)

```powershell
.\query-backend.ps1 -Count -Table "TbClientes"
.\query-backend.ps1 -Count -Table "TbPedidos" -Backend backend_principal
```

---

### 5. Valores únicos de un campo (`-Distinct`)

Útil para explorar dominios de datos o preparar casos de test. Filtra NULLs; incluye cadenas vacías.

```powershell
.\query-backend.ps1 -Distinct -Table "TbClientes" -Field "Estado"
.\query-backend.ps1 -Distinct -Table "TbPedidos" -Field "TipoEstado" -Backend backend_principal
```

---

### 6. Comparar dos backends (`-Compare`)

Ejecuta la misma SQL en dos backends y muestra qué valores están solo en uno de ellos.  
La SQL **debe devolver una sola columna** (normalmente el ID primario).

```powershell
.\query-backend.ps1 -Compare `
    -CompareSQL "SELECT ID FROM TbClientes WHERE F_Baja IS NULL" `
    -Backend backend_principal `
    -CompareBackend backend_cache
```

**Salida:**
```
=== COMPARE ===
SQL   : SELECT ID FROM TbClientes WHERE F_Baja IS NULL
Left  : backend_principal
Right : backend_cache

  backend_principal : 142 filas
  backend_cache     : 139 filas
  RESULT: DIFERENTES
  Solo en backend_principal (3): 301, 302, 303
```

---

## Ejemplos por casos de uso

### Explorar el modelo de datos de un proyecto nuevo

```powershell
# 1. Ver qué tablas tiene el backend
.\query-backend.ps1 -ListTables

# 2. Ver qué tablas son linked (pueden fallar si la red no está disponible)
.\query-backend.ps1 -LinkedTables

# 3. Ver la estructura de una tabla concreta
.\query-backend.ps1 -GetSchema -Table "TbDocumentos"

# 4. Ver cuántos registros tiene
.\query-backend.ps1 -Count -Table "TbDocumentos"

# 5. Ver una muestra de datos
.\query-backend.ps1 -SQL "SELECT TOP 5 * FROM TbDocumentos"
```

---

### Verificar una SQL antes de implementarla en VBA

```powershell
# Probar el JOIN que se va a usar en VBA
.\query-backend.ps1 -SQL "SELECT D.ID, D.Referencia, U.Nombre FROM TbDocumentos D LEFT JOIN TbUsuarios U ON D.IDUsuario = U.ID WHERE D.Estado = 'Borrador'"

# Si falla por tabla linked, identificar cuáles son linked
.\query-backend.ps1 -LinkedTables

# Probar con TOP para no esperar si la tabla es grande
.\query-backend.ps1 -SQL "SELECT * FROM TbDocumentos WHERE Estado = 'Borrador'" -Top 3
```

---

### Preparar casos de test con datos reales

```powershell
# Ver qué valores posibles tiene un campo (para cubrir todos los casos)
.\query-backend.ps1 -Distinct -Table "TbDocumentos" -Field "Estado"

# Buscar un registro con una condición concreta para usar como fixture
.\query-backend.ps1 -SQL "SELECT TOP 1 * FROM TbDocumentos WHERE Estado = 'Cerrado' AND F_Baja IS NOT NULL"

# Buscar IDs que existen en el backend real pero no en la caché
.\query-backend.ps1 -Compare `
    -CompareSQL "SELECT ID FROM TbDocumentos" `
    -Backend backend_principal `
    -CompareBackend backend_cache
```

---

### Diagnosticar una discrepancia entre backend y caché

```powershell
# 1. Comparar IDs activos entre ambos backends
.\query-backend.ps1 -Compare `
    -CompareSQL "SELECT ID FROM TbClientes WHERE F_Baja IS NULL" `
    -Backend backend_principal `
    -CompareBackend backend_cache

# 2. Inspeccionar los registros que sobran en el principal
.\query-backend.ps1 -SQL "SELECT * FROM TbClientes WHERE ID IN (301, 302, 303)" -Backend backend_principal

# 3. Confirmar que no están en la caché
.\query-backend.ps1 -SQL "SELECT * FROM TbClientes WHERE ID IN (301, 302, 303)" -Backend backend_cache
```

---

### Usar con un .accdb externo al proyecto

```powershell
# Sin tocar backends.json, apuntando directo al fichero
.\query-backend.ps1 -ListTables -BackendPath "C:\otro_proyecto\Datos.accdb"
.\query-backend.ps1 -GetSchema -Table "TbProductos" -BackendPath "C:\otro_proyecto\Datos.accdb"
.\query-backend.ps1 -SQL "SELECT * FROM TbProductos" -BackendPath "C:\otro_proyecto\Datos.accdb" -Password "otrapassword"
```

---

## Notas críticas para IA

1. **Access SQL ≠ SQL estándar.** Diferencias clave:
   - Wildcard en LIKE: `*` en lugar de `%` → `WHERE Nombre LIKE "Mar*"`
   - Fechas: `DATE()` para hoy, `DateAdd("m", -3, DATE())` para restar 3 meses
   - Nombres con espacios o caracteres especiales van entre corchetes: `[Mi Tabla]`, `[Mi Campo]`

2. **Tablas linked:** pueden apuntar a rutas de red o a otros `.accdb`. Si un JOIN falla con error de ruta o "tabla no encontrada", la tabla es linked y no está accesible desde el contexto actual. Usar `-LinkedTables` para identificarlas antes de escribir JOINs.

3. **`-Compare` espera una sola columna.** Si la SQL devuelve más de una columna, solo se usa la primera para la comparación. Usar siempre `SELECT ID FROM ...` o similar.

4. **`-Top` solo aplica al modo `-SQL`.** Los modos `-Count`, `-Distinct`, `-GetSchema`, etc. no están limitados por `-Top`.

5. **`-BackendPath` ignora `backends.json` completamente.** Útil para probar un `.accdb` de otro proyecto sin modificar la configuración.

6. **`backends.json` es por proyecto.** Cada proyecto tiene su propio `backends.json` con sus rutas y passwords. No compartir el mismo fichero entre proyectos.

---

## Combinaciones de parámetros y comportamiento ante errores

### Combinaciones requeridas

Algunos modos necesitan parámetros adicionales. Si faltan, el script muestra el help y sale sin ejecutar nada:

| Modo | Parámetros obligatorios | Resultado si faltan |
|------|------------------------|---------------------|
| `-GetSchema` | `-Table` | Muestra help, no ejecuta |
| `-Count` | `-Table` | Muestra help, no ejecuta |
| `-Distinct` | `-Table` + `-Field` | Muestra help, no ejecuta |
| `-Compare` | `-CompareSQL` | Muestra help, no ejecuta |
| `-SQL` | el propio valor de `-SQL` | Muestra help, no ejecuta |

### Errores comunes y su causa

| Error | Causa más probable | Acción |
|-------|--------------------|--------|
| `"El proveedor no está registrado en el equipo local"` | `Microsoft.ACE.OLEDB.12.0` no instalado | Instalar Access Database Engine 2016 |
| `"No es una contraseña válida"` | Password incorrecta o vacía en `backends.json` | Preguntar al usuario la password correcta; no reintentar con variaciones |
| `"No se puede encontrar el archivo"` | Ruta del `.accdb` incorrecta o no accesible | Verificar ruta en `backends.json` o usar `-BackendPath` |
| `"La instrucción SQL no es válida"` | Sintaxis no compatible con Access | Revisar notas de Access SQL: wildcards `*`, fechas, corchetes |
| `"No se puede encontrar la tabla"` en un JOIN | La tabla es linked y su origen no está accesible | Usar `-LinkedTables` para identificarla; evitar ese JOIN o usar la tabla local equivalente |
| `"backends.json no encontrado"` | El script no está en el mismo directorio que `backends.json` | Ejecutar desde el directorio de la skill, o usar `-BackendPath` |
