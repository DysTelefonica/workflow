# SKILL.md — access-localize: Conversión de Tablas Vinculadas Access a Locales

## Objetivo

`access-localize` convierte todas las tablas vinculadas de un frontend Access (`.accdb`) en tablas locales reales, eliminando dependencias de backends externos. El resultado es un frontend **autocontenido** que funciona sin acceso a las bases de datos origen.

Casos de uso:

1. **Distribución offline** — entregar un frontend que funcione sin conexión al servidor/red
2. **Snapshots de datos** — congelar el estado de los datos vinculados en un momento determinado
3. **Migración** — preparar un frontend para moverlo a otro entorno sin arrastrar backends
4. **Testing** — crear una copia funcional aislada del sistema real

> **Nota:** Esta skill opera sobre **frontends** (`.accdb` con tablas vinculadas a otros `.accdb`).
> Para crear sandboxes a partir de backends puros, usar `access-sandbox`.
> Para consultar datos dentro de un `.accdb`, usar `access-query`.

---

## División de responsabilidades

| Skill | Rol | Entrada típica |
|---|---|---|
| `access-localize` | Convierte vinculadas → locales en un frontend | Frontend `.accdb` con vínculos a backends |
| `access-sandbox` | Provisioning: copia, sidecar, localize de backends | Backend `.accdb` protegido |
| `access-query` | Read/query/seed/cleanup sobre un backend existente | Cualquier `.accdb` |

---

## Supuestos

- **Windows** con Microsoft Access instalado (COM `Access.Application` + DAO)
- **PowerShell 5.1+** (o PowerShell 7+)
- Los backends referenciados por las tablas vinculadas deben estar **accesibles** y no bloqueados por otro proceso
- La contraseña de cada backend, si existe, está embebida en `TableDef.Connect` (campo `PWD=`)
- Si en `Connect` no aparece `PWD`, se asume que el backend no tiene contraseña
- Solo se procesan tablas vinculadas a otros ficheros Access (`MS Access;...`). Tablas ODBC, Excel, SharePoint u otros orígenes se ignoran

---

## Capacidades del PS1 (ConvertLinkedAccessTablesToLocal.ps1)

### Parámetros

| Parámetro | Tipo | Requerido | Default | Descripción |
|---|---|---|---|---|
| `-FrontendPath` | `string` | Sí | — | Ruta absoluta al frontend `.accdb` |
| `-FrontendPassword` | `string` | No | `""` | Contraseña del frontend (vacío si no tiene) |
| `-KeepCopiedBackends` | `switch` | No | `$false` | No eliminar los sidecars al finalizar con éxito |
| `-CleanPreviousSidecars` | `switch` | No | `$false` | Eliminar sidecars huérfanos de ejecuciones previas antes de empezar |
| `-BackupFolder` | `string` | No | `""` | Carpeta alternativa para el backup. Si vacío, se crea junto al frontend |

### Contraseñas de backends

No hay parámetro `-BackendPassword`. La contraseña de cada backend se extrae automáticamente de `TableDef.Connect` de las tablas vinculadas. Esta es la fuente de verdad.

---

## Pipeline por tabla (enfoque híbrido DAO + SQL)

Para cada tabla vinculada detectada, el script ejecuta este pipeline atómico:

1. **Abrir definición origen** — abre el sidecar vía DAO, obtiene el `TableDef` de la tabla fuente
2. **Crear tabla local temporal** — `CreateTableDef` campo a campo con DAO, copiando: `Name`, `Type`, `Size`, `Attributes` (incluido AutoNumber), `Required`, `AllowZeroLength`, `DefaultValue`, `ValidationRule`, `ValidationText`, `Description`
3. **Crear vínculo temporal al sidecar** — `TableDef` vinculada temporal apuntando al sidecar con la `Connect` original (reescribiendo `DATABASE=`)
4. **Copiar datos** — `INSERT INTO [TmpLocal] (campos...) SELECT campos... FROM [TmpLink]` con lista explícita de campos (preserva valores AutoNumber)
5. **Validar conteo** — `SELECT COUNT(*)` en origen y destino deben coincidir
6. **Recrear índices y PK** — DAO: `CreateIndex`, copiando `Primary`, `Unique`, `Required`, `IgnoreNulls`, `Clustered` y campos del índice
7. **Sustituir** — eliminar la tabla vinculada original, renombrar la temporal al nombre final
8. **Limpiar** — eliminar el vínculo temporal auxiliar

**Si cualquier paso falla, la tabla no se sustituye** — la vinculada original se mantiene intacta.

---

## Orden exacto del proceso completo

### Fase 0 — Validación
- Resolver y normalizar `FrontendPath`
- Verificar que el fichero existe
- Verificar que DAO está disponible (`New-DaoDbEngine`)
- Detectar sidecars previos (`*__sidecar__*`):
  - Si existen y no se pasa `-CleanPreviousSidecars` → **abortar**
  - Si existen y se pasa `-CleanPreviousSidecars` → eliminarlos

### Fase 1 — Backup
- Crear backup del frontend: `NombreFrontend__backup__yyyyMMdd_HHmmss.accdb`
- Si falla el backup → **abortar** (no se modifica nada)

### Fase 2 — Apertura controlada
- `AllowBypassKey`: leer estado original → habilitar temporalmente
- `AutoExec` / `StartupForm`: deshabilitar temporalmente (renombra macro, elimina propiedad)
- Crear instancia COM: `Access.Application` con `Visible=$false`, `UserControl=$false`, `AutomationSecurity=1`
- `OpenCurrentDatabase` con contraseña
- `DoCmd.SetWarnings($false)`
- Capturar PID de Access (P/Invoke `GetWindowThreadProcessId`)

### Fase 3 — Descubrimiento
- Leer `CurrentDb.TableDefs`
- Filtrar: excluir `MSys*`, excluir locales, incluir solo `Connect` tipo `MS Access;...`
- Agrupar por `DATABASE=` (ruta del backend)

### Fase 4 — Copia de sidecars
- Por cada backend: verificar existencia → verificar accesibilidad → copiar al lado del frontend → verificar accesibilidad del sidecar

### Fase 5 — Procesamiento
- Por cada backend → por cada tabla vinculada: ejecutar el pipeline atómico (sección anterior)

### Fase 6 — Cierre y restauración
- Cerrar base y Access COM
- Restaurar `AllowBypassKey`, `AutoExec`, `StartupForm`
- Liberar COM + `GC.Collect()`
- Kill del PID de Access como red de seguridad
- En éxito: eliminar sidecars (salvo `-KeepCopiedBackends`)
- En error: conservar sidecars para diagnóstico

---

## Estructura del skill

```
.agents/skills/access-localize/
├── SKILL.md                                  # Este archivo (especificación)
├── ConvertLinkedAccessTablesToLocal.ps1       # Motor PowerShell
├── README.md                                 # Guía de uso rápida + ejemplos
```

---

## Invocación

```powershell
# Caso básico — frontend con password, backends con password embebida en Connect
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd"

# Con backup en carpeta separada
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd" `
    -BackupFolder "C:\backups"

# Limpiar sidecars de una ejecución previa fallida
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd" `
    -CleanPreviousSidecars

# Conservar sidecars para inspección
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd" `
    -KeepCopiedBackends
```

---

## Preservación de estructura

### Qué se preserva
- Tipo de campo (incluido AutoNumber con valores originales)
- Tamaño de campo
- Atributos (`dbAutoIncrField`, `dbFixedField`, etc.)
- Required, AllowZeroLength
- DefaultValue, ValidationRule, ValidationText
- Description (propiedad custom)
- Índices: Primary, Unique, Required, IgnoreNulls, Clustered, campos y orden

### Qué NO se preserva en esta versión
- Relaciones (Relations) entre tablas
- Propiedades avanzadas no estables o dependientes de contexto
- Tablas vinculadas a orígenes no-Access (ODBC, SQL Server, Excel, SharePoint)

---

## Gestión de errores

### Error fatal (aborta todo)
- Frontend no existe
- Backup falla
- DAO no disponible
- No se puede abrir frontend
- No se puede desactivar AutoExec/StartupForm
- Backend referenciado no existe
- Sidecar no se puede copiar o abrir
- Tabla temporal no se puede crear
- Índice esencial falla al recrearse
- Conteo de registros no coincide

### Error no fatal (se registra y continúa)
- Propiedad de campo individual no portable (Description fallida, ValidationRule incompatible)
- Propiedad de índice secundaria no copiable

---

## Relación con access-sandbox y access-query

| Escenario | Skill a usar |
|---|---|
| Tengo un frontend con vínculos y quiero hacerlo autónomo | `access-localize` |
| Tengo un backend protegido y quiero crear un sandbox sin vínculos | `access-sandbox` |
| Quiero consultar/modificar datos en un `.accdb` | `access-query` |
| Quiero generar un ERD desde un frontend ya localizado | `access-query` (sobre el resultado de `access-localize`) |

**Flujo combinado típico:**
1. `access-localize` convierte el frontend en autónomo
2. `access-query` opera sobre el frontend ya localizado

---

## Casos límite

- **Frontend sin tablas vinculadas** → el script termina limpiamente informando "Nada que hacer"
- **Tablas vinculadas a orígenes no-Access** → se ignoran silenciosamente (solo se procesan `MS Access;...`)
- **Backend bloqueado por otro proceso** → error al copiar sidecar (error claro con ruta)
- **Sidecars huérfanos de ejecución previa** → aborta salvo `-CleanPreviousSidecars`
- **Tablas con mismo nombre en distintos backends** → cada una se procesa independientemente
- **Frontend ya localizado (sin vínculos)** → termina con "Nada que hacer"

---

## Requisitos

- Windows con Microsoft Access instalado (no basta Access Database Engine — necesita `Access.Application` COM)
- PowerShell 5.1+ o PowerShell 7+
- Permisos de lectura en los backends origen
- Permisos de escritura en la carpeta del frontend (para sidecars y backup)
