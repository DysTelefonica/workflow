# access-localize — Conversión de Tablas Vinculadas Access a Locales

Convierte todas las tablas vinculadas de un frontend Access en tablas locales reales,
eliminando dependencias de backends externos. El frontend queda autocontenido.

> **Rol:** Esta skill *localiza* un frontend. Para crear sandboxes desde backends, usar **`access-sandbox`**.
> Para consultar o modificar datos, usar **`access-query`**.

---

## Archivos

| Archivo | Descripción |
|---|---|
| `ConvertLinkedAccessTablesToLocal.ps1` | Motor de conversión (PowerShell) |
| `SKILL.md` | Especificación técnica completa (referencia canónica) |
| `README.md` | Este archivo — guía de uso rápida |

---

## Quick start

```powershell
# Caso típico — frontend con password, backends con password en Connect
& "C:\Users\adm\.config\opencode\.agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd"
```

**Resultado:** Todas las tablas que antes eran vínculos a backends externos son ahora tablas locales con datos copiados. El frontend funciona sin acceso a los backends.

---

## Parámetros

| Parámetro | Default | Descripción |
|---|---|---|
| `-FrontendPath` | (requerido) | Ruta absoluta al frontend `.accdb` |
| `-FrontendPassword` | `""` | Contraseña del frontend |
| `-KeepCopiedBackends` | `$false` | No eliminar sidecars al finalizar |
| `-CleanPreviousSidecars` | `$false` | Limpiar sidecars huérfanos antes de empezar |
| `-BackupFolder` | `""` | Carpeta alternativa para el backup |

> **No hay `-BackendPassword`**: la contraseña de cada backend se extrae automáticamente de
> `TableDef.Connect` de las tablas vinculadas.

---

## Qué hace internamente

```
Frontend.accdb                      Backend_A.accdb    Backend_B.accdb
┌──────────────────┐                ┌──────────────┐   ┌──────────────┐
│ tblClientes ──────────────────────│► tblClientes │   │              │
│ tblPedidos ───────────────────────│► tblPedidos  │   │              │
│ tblProductos ─────────────────────│──────────────│───│► tblProductos│
└──────────────────┘                └──────────────┘   └──────────────┘

                            ▼  access-localize  ▼

Frontend.accdb (localizado)
┌──────────────────┐
│ tblClientes      │  ← tabla local con datos copiados
│ tblPedidos       │  ← tabla local con datos copiados
│ tblProductos     │  ← tabla local con datos copiados
└──────────────────┘
   Sin vínculos externos
```

### Pipeline por tabla

1. Abre backend vía DAO → lee definición de la tabla origen
2. Crea tabla local temporal campo a campo (preserva tipos, AutoNumber, propiedades)
3. Crea vínculo temporal al sidecar
4. `INSERT INTO ... SELECT ...` con lista explícita de campos (preserva AutoNumber)
5. Valida conteo de registros
6. Recrea índices y PK desde el origen
7. Elimina la vinculada original → renombra la temporal al nombre final

---

## Ejemplos

### Caso básico

```powershell
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd"
```

### Con backup en carpeta separada

```powershell
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd" `
    -BackupFolder "D:\backups\access"
```

### Limpiar sidecars de ejecución previa fallida

```powershell
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd" `
    -CleanPreviousSidecars
```

### Conservar sidecars para inspección

```powershell
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd" `
    -KeepCopiedBackends
```

---

## Seguridad y consistencia

- **Backup obligatorio** antes de cualquier modificación. Nombre: `Frontend__backup__yyyyMMdd_HHmmss.accdb`
- **Sustitución atómica**: la tabla vinculada original solo se elimina cuando la local temporal está completa y validada
- **Proceso oculto**: `Access.Application` con `Visible=$false`, `UserControl=$false`, `AutomationSecurity=1`
- **AutoExec/StartupForm deshabilitados** temporalmente durante la operación y restaurados al final
- **Kill de seguridad**: si el proceso COM de Access no se cierra limpiamente, se mata por PID

---

## Preservación de estructura

| Se preserva | No se preserva (esta versión) |
|---|---|
| Tipos de campo (incluido AutoNumber con valores) | Relaciones (Relations) entre tablas |
| Tamaño, Attributes, Required, AllowZeroLength | Propiedades avanzadas no portables |
| DefaultValue, ValidationRule, ValidationText | Vínculos a orígenes no-Access |
| Description | |
| Índices: PK, Unique, IgnoreNulls, Clustered | |

---

## Gestión de errores

| Tipo | Comportamiento |
|---|---|
| Frontend no existe / backup falla / DAO no disponible | **Aborta** — no se modifica nada |
| Backend referenciado no accesible | **Aborta** — backup disponible |
| Tabla individual falla (creación/datos/índices) | **Aborta** — tablas ya procesadas quedan bien, las no procesadas quedan como vínculos |
| Propiedad de campo individual no portable | **Continúa** — se registra y se salta |

En caso de error, el backup siempre está disponible y los sidecars se conservan para diagnóstico.

---

## Flujo combinado con otras skills

```powershell
# 1. Localizar el frontend (esta skill)
& ".agents\skills\access-localize\ConvertLinkedAccessTablesToLocal.ps1" `
    -FrontendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -FrontendPassword "dpddpd"

# 2. Consultar datos del frontend ya localizado (access-query)
& ".agents\skills\access-query\query-backend.ps1" `
    -BackendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -Password "dpddpd" `
    -SQL "SELECT COUNT(*) FROM tblClientes"

# 3. Generar ERD desde el frontend localizado
& ".agents\skills\access-query\query-backend.ps1" `
    -BackendPath "C:\proyecto\MiApp_Gestion.accdb" `
    -Password "dpddpd" `
    -ListTables
```

---

## Notas de implementación

- **Enfoque híbrido**: DAO para estructura + `INSERT INTO...SELECT` para datos (no usa `SELECT INTO` ni recordset fila a fila)
- **AutoNumber**: se preservan valores originales incluyendo el campo explícitamente en `INSERT INTO...SELECT`
- **COM cleanup**: `FinalReleaseComObject` + una sola pasada de `GC.Collect()` + `GaitForPendingFinalizers()` en los puntos de cierre
- **DAO versions**: intenta 160 → 150 → 140 → 120 → 36 en cascada
- **Sidecars**: copias temporales de los backends junto al frontend, nombradas `Frontend__sidecar__Backend.accdb`
- **P/Invoke**: usa `GetWindowThreadProcessId` para capturar el PID de `MSACCESS.EXE` y poder matarlo si no se cierra limpiamente

---

## Requisitos

- Windows con Microsoft Access instalado (necesita `Access.Application` COM, no basta Access Database Engine)
- PowerShell 5.1+ o PowerShell 7+
- Permisos de lectura en los backends origen
- Permisos de escritura en la carpeta del frontend
