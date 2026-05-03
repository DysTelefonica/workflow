# SKILL.md — Skill para workflow Access/VBA (Export → Trabajo → Sync → Compilar → ERD → Cierre)

## Objetivo
Definir un **skill** que automatice el workflow de desarrollo y documentación en un proyecto Microsoft Access/VBA:

1) **Al inicio** de una nueva feature/fix: **Exportar TODOS los módulos** del proyecto VBA a disco (snapshot base).
2) Se trabaja sobre la mejora editando los archivos exportados (normalmente con IA).
3) **Todo módulo modificado por la IA debe sincronizarse** (Import) hacia el VBA real de la BD.
4) Tras cada sincronización, el skill **debe proponer al usuario compilar** el proyecto en el VBE.
5) **Generación de documentación**: Extraer estructura de tablas (ERD/Diccionario) a Markdown para contexto de la IA.
6) **Al cerrar** la tarea (fin de sesión): export final opcional (snapshot consistente) + resumen.

> El skill es **autocontenido**: incluye `VBAManager.ps1` y todo lo necesario para ejecutarse sin dependencias externas.

---

## Alcance y supuestos
- Entorno: **Windows** con Microsoft Access instalado (automatización COM y DAO).
- El repositorio contiene una BD Access (`.accdb/.accde/.mdb/.mde`) en la raíz del proyecto, o el usuario la pasa por parámetro.
- La exportación se guarda bajo una carpeta configurable (default `src/`).
- La documentación se genera en `docs/ERD/` o ruta configurable.
- El skill cierra automáticamente cualquier instancia de Access que tenga abierta la BD objetivo (usando ROT para no afectar otras BDs abiertas).

---

## Capacidades del PS1 (VBAManager.ps1)

El PS1 es el motor que ejecuta todas las operaciones sobre Access. Acepta los siguientes parámetros:

| Parámetro | Tipo | Descripción |
|---|---|---|
| `-Action` | `Export\|Import\|Fix-Encoding\|Generate-ERD\|List-Objects\|Exists\|Run-Procedure\|Compile` | **Obligatorio**. Acción a ejecutar |
| `-AccessPath` | string | Ruta a la BD frontend |
| `-Password` | string | Contraseña de la BD (default en el script) |
| `-DestinationRoot` | string | Carpeta raíz de export/import (default: `src`) |
| `-ModuleName` | string[] | Uno o más nombres de módulo para operaciones selectivas |
| `-ModuleNamesJson` | string | Transporte interno usado por `handler.js` para listas de módulos; evita problemas con comas y binding posicional de PowerShell |
| `-ProcedureName` | string | Nombre de función/sub pública para `Run-Procedure` |
| `-ProcedureArgsJson` | string | Array JSON de argumentos simples para `Run-Procedure` |
| `-ImportMode` | `Auto\|Form\|Code` | Solo para `Import`: `Auto` (default), `Form` (forzar `.form.txt/.frm`), `Code` (forzar `.cls/.bas`) |
| `-BackendPath` | string | Ruta al backend `*_Datos.accdb` para Generate-ERD; si falta, se autodetecta |
| `-ErdPath` | string | Carpeta de salida del ERD |
| `-Location` | `Both\|Src\|Access` | Ámbito de Fix-Encoding (default: `Both`) |
| `-Json` | switch | Salida JSON para `List-Objects`, `Exists` y `Run-Procedure` |

### Acciones

**`Export`** — Exporta módulos VBA de la BD a disco:
- Sin `-ModuleName`: exporta todos los módulos
- Con `-ModuleName A B C`: exporta solo los indicados
- Formularios: usa `SaveAsText` → `.form.txt` (UI + código) y `.cls` (solo código)
- Módulos/clases: usa `VBComponents.Export` → `.bas` / `.cls`

**`Import`** — Importa módulos desde disco a la BD:
- Sin `-ModuleName`: importa todos los archivos de `src/`
- Con `-ModuleName A B C`: importa solo los indicados
- Modo por defecto `-ImportMode Auto`: prioridad `.form.txt` > `.frm` > `.cls` > `.bas`
- Si existe `.form.txt` de un formulario, `Auto` lo trata como **documento completo** y lo importa completo aunque el cambio original haya sido solo de código
- Si falla la reconstrucción del header canónico desde Access durante ese import, el proceso **aborta** para no continuar con un header local potencialmente stale
- Para módulos/clases **nuevos**, la skill intenta primero clonarlos desde un componente persistido del mismo tipo y solo cae al alta desde cero si no existe ninguna semilla disponible
- Si para un mismo formulario/reporte existen dos sidecars `.cls` gemelos (`Form_X.cls` y `X.cls`) y divergen, la skill prioriza el más recientemente modificado, unifica el gemelo restante y luego sincroniza el `CodeBehind*` del `.form.txt`
- `-ImportMode Form`: importa solo layout/formulario (`.form.txt/.frm`)
- `-ImportMode Code`: importa solo code-behind (`.cls/.bas`)
- Formularios (`.form.txt`): usa `LoadFromText` — completamente silencioso
- Módulos/clases: usa `DeleteLines + AddFromFile` — sin diálogos VBE

**`Fix-Encoding`** — Corrige encoding ANSI↔UTF-8:
- `-Location Src`: corrige BOM en archivos de `src/`
- `-Location Access`: reimporta desde `src/` para corregir en la BD
- `-Location Both` (default): ambos
- Selectivo con `-ModuleName`

**`Generate-ERD`** — Genera ERD en Markdown:
- Lee tablas, campos, tipos DAO, PKs e índices, y relaciones
- Detecta backends vinculados no alcanzables y los documenta
- Autodetecta `*_Datos.accdb` si no se pasa `-BackendPath`
- No requiere `-AccessPath`: trabaja contra `-BackendPath` o contra el backend autodetectado

**`List-Objects`** — Lista los objetos reales del frontend:
- forms
- reports
- modules
- classes
- documentModules

**`Exists`** — Inspecciona un nombre concreto y devuelve:
- si existe como objeto Access
- si existe como VBComponent
- si es document module
- el nombre real resuelto
- el modo recomendado (`import`)

---

## Comandos del CLI (cli.js)

El CLI es la interfaz de usuario sobre el PS1. Todos los comandos que expone mapean directamente a acciones del PS1.

### Modelo operativo actual

La skill tiene **dos modos**:

1. **Modo canónico / on-demand (stateless)**  
   Usar para casi todo:
   - `import`
   - `export`
   - `exists`
   - `list-objects`
   - `fix-encoding`
   - `generate-erd`

   Estos comandos deben funcionar por sí solos, usando los flags actuales (`--access`, `--destination_root`) y **sin depender de una sesión previa**.

2. **Modo sesión / watch (opcional)**  
   Solo para workflows largos con auto-sync:
   - `start`
   - `watch`
   - `status`
   - `end`

   Este modo mantiene estado en `.access-vba-skill/session.json` y no debe considerarse el camino normal para imports manuales.

### Tabla de comandos

| Comando CLI | Acción PS1 | Módulos | Descripción |
|---|---|---|---|
| `start` | `Export` | todos | Export inicial + inicia sesión (modo watch; opcional) |
| `watch` | `Import` (auto) | modificados | Auto-sync al detectar cambios en src/ (modo sesión) |
| `export <Mod...>` | `Export` | selectivo | Exporta módulos específicos |
| `export-all` | `Export` | todos | Exporta todos sin iniciar sesión |
| `import <Mod...>` | `Import` | selectivo | **Comando canónico**: detecta si cada entrada es módulo, clase o formulario y hace el import correcto |
| `import-form <Mod...>` | `Import` (`ImportMode=Form`) | selectivo | Importa formularios desde `*.form.txt`/`*.frm` (UI + código) — uso avanzado |
| `import-code <Mod...>` | `Import` (`ImportMode=Code`) | selectivo | Importa solo code-behind desde `*.cls`/`*.bas`; bloquea crear `Módulo1`/`Módulo2` si el target parece formulario/reporte — uso avanzado |
| `import-all` | `Import` | todos | Importa todo src/ |
| `list-objects` | `List-Objects` | — | Lista los objetos reales del frontend; ideal para diagnóstico |
| `exists <Mod>` | `Exists` | uno | Verifica si un nombre existe realmente en Access/VBA y cómo se resolvió |
| `compile-vba` | `Compile` | — | Gate de compilación: ejecuta `acCmdCompileAndSaveAllModules` y devuelve fase/ubicación si falla |
| `run-vba <Proc>` | `Run-Procedure` | — | Ejecuta una función/sub pública con `Access.Application.Run`; acepta `--args-json` |
| `test-vba` | `Compile` + `Run-Procedure` | — | Runner externo TDD: compila y ejecuta un plan JSON de tests |
| `sync <Mod...>` | `Import` | selectivo | Alias de import |
| `fix-encoding [Mod...]` | `Fix-Encoding` | selectivo/todos | Corrige encoding |
| `generate-erd` | `Generate-ERD` | — | Genera ERD Markdown |
| `verify-code [Mod...]` | — | selectivo/todos | Diffa `.cls` vs CodeBehind del `.form.txt` sin importar. Sin módulos: todos los formularios/reportes. Exit 1 si hay diferencias. |
| `status` | — | — | Muestra estado de sesión/watch |
| `end` | `Export` (opcional) | todos | Cierra sesión/watch + export final |

`import-all` puede hacer varias pasadas internas si detecta fallos por orden de dependencias entre módulos/clases. Esas pasadas las gestiona `VBAManager.ps1` dentro de la acción `Import`; `handler.js` lanza una sola llamada protegida por backup/restore. Si aun así quedan pendientes, termina con error agregado y no reporta falso OK.

### Ejecución secuencial obligatoria

⚠️ **Cada comando `import`, `import-code` o `import-form` cierra Access al inicio** (Kill via ROT). Por eso:

- **NUNCA encadenar comandos con `&&` o `;`** — el segundo comando falla porque Access ya fue cerrado por el primero
- **Ejecutar cada comando en su PROPIA llamada** — esperar el resultado antes de invocar el siguiente
- **Preferir siempre `import`** como comando único, incluso si la lista mezcla módulos, clases y formularios

```powershell
# ✅ Correcto: un solo import canónico, incluso heterogéneo
node cli.js import NombreModulo ClaseServicio subfrmDatosPCSUB_DictamenRAC --access "MiBD.accdb"

# ❌ Incorrecto: falla en el segundo comando
node cli.js import NombreModulo --access "MiBD.accdb" && node cli.js import-code Form_A --access "MiBD.accdb"
```

### Flags del CLI

| Flag | Mapea a PS1 | Default |
|---|---|---|
| `--access <ruta>` | `-AccessPath` | Autodetecta en CWD; la BD debe estar en el directorio desde donde ejecutás el CLI |
| `--password <pwd>` | `-Password` | `ACCESS_VBA_PASSWORD` env var recomendado (ver abajo) |
| `--destination_root <dir>` / `--destination <dir>` | `-DestinationRoot` | `src` |
| `--location Both\|Src\|Access` | `-Location` | `Both` |
| `--backend <ruta>` | `-BackendPath` | Autodetecta `*_Datos.accdb` |
| `--erd_path <dir>` | `-ErdPath` | `docs/ERD` |
| `--args-json <json-array>` | `-ProcedureArgsJson` | `[]`; solo para `run-vba` |
| `--tests <ruta>` | — | `tests.vba.json`; solo para `test-vba` |
| `--procedure <Proc>` / `test-vba <Proc>` | `-ProcedureName` | Ejecuta un procedimiento público como test directo, sin `tests.vba.json` |
| `--filter <texto>` | — | filtra `test-vba` por nombre/procedimiento/tag |
| `--no-compile` | — | salta el compile gate de `test-vba` |
| `--allow-startup-execution` | `-AllowStartupExecution` | **Avanzado/peligroso**: abre aunque no pueda deshabilitar AutoExec/StartupForm |
| `--json` | `-Json` | `false` |
| `--debounce_ms <n>` | — (Node.js) | `600` |
| `--auto_export_on_end false` | — (Node.js) | `true` |

### Contraseña de la BD

La contraseña se resuelve en este orden de prioridad:

1. `--password <pwd>` en el comando (no recomendado: queda expuesto en el historial de proceso)
2. Variable de entorno `ACCESS_VBA_PASSWORD` (**forma recomendada**)
3. Sin contraseña (BDs no protegidas)

```powershell
# Configurar una vez en la sesión (o en el perfil de PowerShell / .env del proyecto)
$env:ACCESS_VBA_PASSWORD = "miPassword"

# A partir de aquí, todos los comandos la usan sin pasarla explícitamente
node cli.js list-objects --access "CONDOR.accdb" --json
node cli.js import Form_frmGestion --access "CONDOR.accdb"
```

---

## Introspección del frontend

Cuando una IA dude de si un formulario, subformulario, reporte o módulo existe realmente en el binario, **no debe adivinar**. Debe inspeccionar el frontend.

### Listado completo

```powershell
node cli.js list-objects --access "CONDOR.accdb" --json
```

### Verificación puntual

```powershell
node cli.js exists subfrmDatosPCSUB_DictamenRAC --access "CONDOR.accdb" --json
```

Uso recomendado:
- si falla un import de formulario/reporte
- si hay duda entre nombre Access (`subfrmX`) y document module VBA (`Form_subfrmX`)
- si la IA no sabe si el target existe o no en la BD

---

## Ejecución de funciones/subs públicas desde la IA

La primitiva para que una IA valide lo que acaba de programar en VBA es `run-vba`.

```powershell
# Sin argumentos
node cli.js run-vba NombrePublico --access "CONDOR.accdb" --json

# Con argumentos simples
node cli.js run-vba NombrePublico --access "CONDOR.accdb" --args-json "[123, \"texto\", true, null]" --json
```

Contrato:
- `NombrePublico` debe ser una `Public Function` o `Public Sub` accesible vía `Access.Application.Run`.
- `--args-json` debe ser **un array JSON**. Si se omite, se ejecuta sin argumentos.
- Tipos soportados en argumentos: `string`, `number`, `boolean`, `null`.
- No pasar objetos/arrays anidados al principio: COM + VBA + PowerShell se vuelve frágil. Si necesitás estructura, pasá un `String` JSON y parsealo en VBA.
- `--json` es la forma recomendada para IA: devuelve `ok`, `procedure`, `argsCount`, `returnValue`, `returnType`, `payload`, `logs`, `error`.
- Si `ok=false`, el CLI termina con exit code `1`.
- `Debug.Print` no se captura de forma fiable por COM. Para logs evaluables por IA, la función VBA debe devolver un JSON string con `logs`.
- Si `returnValue` es un JSON string con `{ "ok": false }`, la salida normalizada de `run-vba` también tendrá `ok=false`.

Salida JSON esperada:

```json
{
  "ok": true,
  "procedure": "NombrePublico",
  "argsCount": 2,
  "returnValue": "{\"ok\":true,\"logs\":[\"preparando\",\"resultado OK\"],\"value\":\"OK\"}",
  "returnType": "System.String",
  "payload": {
    "ok": true,
    "logs": ["preparando", "resultado OK"],
    "value": "OK"
  },
  "logs": ["preparando", "resultado OK"],
  "error": null
}
```

Patrón recomendado para tests VBA:

```vb
Public Function Test_CalculaTotal(ByVal entrada As Long) As String
    ' Devolver JSON string; la IA leerá payload/logs normalizados por run-vba.
    Test_CalculaTotal = "{""ok"":true,""logs"":[""preparando"", ""resultado OK""],""value"":42}"
End Function
```

Flujo para IA:

1. Modificar `.bas`/`.cls`/`.form.txt`.
2. Importar con `node cli.js import <Mod...> --access "CONDOR.accdb"`.
3. Ejecutar funciones públicas puntuales con `run-vba`.
4. Si falla, corregir código, reimportar y repetir.

---

## TDD externo para VBA Access

Para TDD con IA, la secuencia correcta es **import → compile gate → tests**. No mezcles errores de compilación con assertions: son fases distintas.

Esta skill no intenta convertir VBA en un ecosistema moderno de testing. Hace algo más práctico para Access: usa Access como runtime real, pero deja el **orquestador de tests fuera** en JSON para que la IA pueda leer, decidir y repetir.

### Regla de oro

Un test VBA público debe ser una función pequeña que:

1. Prepara datos mínimos.
2. Llama a la función real que querés validar.
3. Devuelve un **JSON string** con el resultado.
4. No depende de `Debug.Print`, popups, formularios visibles ni interacción humana.

Contrato recomendado de retorno:

```json
{
  "ok": true,
  "value": 42,
  "logs": ["Arrange listo", "Act ejecutado", "Assert OK"],
  "error": null
}
```

Si el test espera un fallo funcional, devolvé `ok:false` y un `error` semántico:

```json
{
  "ok": false,
  "error": "cliente no existe",
  "logs": ["cliente CLI999 no encontrado"]
}
```

IMPORTANTE: un `ok:false` de negocio no es lo mismo que un error de compilación. Si no compila, `test-vba` corta antes de ejecutar cualquier test.

### Dónde poner los tests en proyectos Access

Patrón recomendado:

```text
src/
  modules/
    Clientes.bas
    Test_Clientes.bas
  classes/
    ClienteService.cls
tests.vba.json
```

Convenciones:

- Módulos de test: `Test_<Área>.bas`.
- Funciones de test: `Public Function Test_<Caso>() As String`.
- Cada test devuelve `String` con JSON.
- El código de producción no debe depender de los módulos `Test_*`.
- Los tests pueden llamar a funciones públicas, clases públicas o wrappers públicos que expongan lógica privada.
- Si necesitás fixtures complejas, preferí helpers públicos `TestSetup_*` / `TestTeardown_*` dentro de módulos de test.

NO uses formularios como runner de tests. Los formularios son UI; el runner externo debe llamar funciones públicas puras o semi-puras.

### Compile gate

```powershell
node cli.js compile-vba --access "CONDOR.accdb" --json
```

Salida:

```json
{
  "ok": false,
  "phase": "compile",
  "error": "Argument not optional",
  "component": "ModuloClientes",
  "line": 42,
  "column": 13,
  "endLine": 42,
  "endColumn": 33,
  "sourceLine": "resultado = CalcularTotal(cliente)"
}
```

Notas:
- Internamente usa `Access.Application.RunCommand(126)` (`acCmdCompileAndSaveAllModules`).
- Si Access/VBE selecciona la línea problemática, la skill intenta devolver `component`, `line`, `column` y `sourceLine`.
- Microsoft no expone un compilador CLI moderno para Access VBA; el texto exacto del popup puede no estar siempre disponible por COM.

Lectura para IA:

- `ok:true` → podés ejecutar tests.
- `ok:false` + `phase:"compile"` → NO intentes interpretar assertions; corregí primero el código que no compila.
- Si hay `component/line/sourceLine`, editá ese módulo y repetí `import → compile-vba`.

### Plan de tests externo

Archivo `tests.vba.json`:

```json
{
  "tests": [
    {
      "name": "calcula total básico",
      "procedure": "Test_CalculaTotal",
      "args": [10],
      "expect": { "ok": true, "value": 42 },
      "tags": ["total"]
    },
    {
      "name": "cliente inexistente",
      "procedure": "Test_ClienteNoExiste",
      "args": ["CLI999"],
      "expect": { "ok": false, "errorContains": "no existe" },
      "tags": ["clientes"]
    }
  ]
}
```

Schema práctico:

```json
{
  "tests": [
    {
      "name": "nombre humano del caso",
      "procedure": "NombrePublicoVba",
      "args": ["string", 123, true, null],
      "expect": {
        "ok": true,
        "value": "resultado esperado",
        "payloadContains": { "campo": "valor" },
        "pathEquals": { "payload.total": 42 }
      },
      "tags": ["area", "rapido"]
    }
  ]
}
```

Campos:

| Campo | Obligatorio | Uso |
|---|---:|---|
| `name` | No | Nombre legible del test; si falta, usa `procedure` |
| `procedure` / `proc` | Sí | Función/sub pública ejecutable por `Access.Application.Run` |
| `args` | No | Array JSON de argumentos simples; default `[]` |
| `expect` | No | Assertions externas que evalúa Node |
| `tags` | No | Etiquetas para `--filter` |

Ejecutar:

```powershell
node cli.js test-vba --access "CONDOR.accdb" --tests tests.vba.json --json
```

Variantes:

```powershell
node cli.js test-vba --filter clientes --json
node cli.js test-vba --no-compile --json

# Sin tests.vba.json: ejecutar un procedimiento público directamente
node cli.js test-vba Canonical_RunAll --json
node cli.js test-vba --procedure Canonical_RunAll --json
```

Contrato de `test-vba`:
- Primero ejecuta `compile-vba`. Si no compila, devuelve `phase: "compile"` y **no ejecuta tests**.
- Si compila, ejecuta cada `procedure` con `run-vba`.
- Si se pasa `test-vba <Proc>` o `--procedure <Proc>`, crea un plan efímero de un solo test y **no exige `tests.vba.json`**.
- Si se usa `--filter` y ningún test coincide, el resultado será `ok: true` con `total: 0`. **Una IA no debe interpretar esto como éxito** sino como que no se encontraron tests para ejecutar.
- Las expectations soportadas son:
  - `ok`
  - `value` (lee `payload.value` o, si no hay payload, `returnValue`)
  - `returnValue`
  - `payloadContains`
  - `errorContains`
  - `pathEquals` para rutas tipo `payload.total`
- Si algún test falla, `test-vba` devuelve `ok:false` y exit code `1`.

### Expectations disponibles

`test-vba` evalúa el JSON normalizado que sale de `run-vba`.

| Expectation | Qué compara | Ejemplo |
|---|---|---|
| `ok` | `result.ok` normalizado. Si se omite, se verifica implícitamente `ok: true`. | `"expect": { "ok": true }` |
| `value` | `payload.value`; si no hay payload, `returnValue` | `"expect": { "value": 42 }` |
| `returnValue` | retorno bruto de VBA | `"expect": { "returnValue": "OK" }` |
| `payloadContains` | subconjunto profundo de `payload` | `"expect": { "payloadContains": { "estado": "OK" } }` |
| `errorContains` | substring dentro de `error` o `payload.error` | `"expect": { "errorContains": "no existe" }` |
| `pathEquals` | paths puntuales con punto | `"expect": { "pathEquals": { "payload.total": 42 } }` |

Regla práctica: para tests nuevos usá `ok` + `value` o `ok` + `payloadContains`. Reservá `returnValue` para funciones legacy que no devuelven JSON.

### Integración con harness existentes

Si el proyecto ya tiene una batería propia, NO la reescribas de entrada. Adaptá de menor a mayor:

1. Si ya existe un procedimiento público tipo `Canonical_RunAll`, ejecutalo directo:
   ```powershell
   node cli.js test-vba Canonical_RunAll --access "CONDOR.accdb" --password "dpddpd" --json
   ```
2. Si el harness expone solo `Private Sub/Function`, creá un wrapper público mínimo:
   ```vb
   Public Function Test_Canonical_RunAll() As String
       On Error GoTo EH
       Canonical_Setup
       Canonical_RunAll
       Canonical_TearDown
       Test_Canonical_RunAll = "{""ok"":true,""logs"":[""Canonical_RunAll OK""]}"
       Exit Function
   EH:
       On Error Resume Next
       Canonical_TearDown
       Test_Canonical_RunAll = "{""ok"":false,""error"":""" & Replace(Err.Description, """", "'") & """}"
   End Function
   ```
3. Si necesitás slices, creá wrappers públicos por slice (`Test_Canonical_DC`, `Test_Canonical_VTS`, etc.) y opcionalmente un `tests.vba.json`.

Regla: `Access.Application.Run` solo puede llamar procedimientos accesibles públicamente desde Access. Si algo es `Private`, la skill no debe intentar invocarlo por magia: envolvelo.

### AutoExec/StartupForm bloqueando la apertura

Por defecto la skill intenta deshabilitar `AutoExec` y `StartupForm` antes de abrir Access. Si falla, aborta para evitar ejecución no desatendida.

En un entorno controlado de testing/CI podés permitir la apertura igualmente:

```powershell
node cli.js test-vba Canonical_RunAll --access "CONDOR.accdb" --password "dpddpd" --allow-startup-execution --json
node cli.js list-objects --access "CONDOR.accdb" --password "dpddpd" --allow-startup-execution --json
```

Usalo solo si aceptás que se ejecute startup code real del frontend. Si el startup muestra UI modal, abre formularios o pide input, el runner puede bloquearse. La solución robusta de proyecto sigue siendo corregir el binario para que el modo automatizado pueda deshabilitar Startup/AutoExec.

### Descubrimiento automático de tests

La skill **no autoejecuta tests descubiertos por nombre**. Puede parecer cómodo, pero en Access real es peligroso: podés ejecutar helpers destructivos o tests que requieren fixtures especiales.

Política:

- Para baterías existentes: `test-vba <Proc>` o `--procedure`.
- Para suites nuevas y repetibles: `tests.vba.json`.
- Para descubrimiento futuro: solo debería añadirse como modo explícito y filtrado, nunca como default.

Salida si falla la compilación:

```json
{
  "ok": false,
  "phase": "compile",
  "testsPath": "tests.vba.json",
  "total": 2,
  "passed": 0,
  "failed": 0,
  "skipped": 2,
  "compile": {
    "ok": false,
    "phase": "compile",
    "error": "Argument not optional",
    "component": "ModuloClientes",
    "line": 42,
    "column": 13,
    "endLine": 42,
    "endColumn": 33,
    "sourceLine": "resultado = CalcularTotal(cliente)"
  },
  "results": []
}
```

Salida agregada (si compila y ejecuta tests):

```json
{
  "ok": false,
  "phase": "tests",
  "testsPath": "tests.vba.json",
  "total": 2,
  "passed": 1,
  "failed": 1,
  "skipped": 0,
  "results": [
    {
      "name": "calcula total básico",
      "procedure": "Test_CalculaTotal",
      "args": [10],
      "ok": true,
      "failures": [],
      "run": {
        "ok": true,
        "logs": ["preparando", "resultado OK"],
        "payload": null
      },
      "logs": ["preparando", "resultado OK"]
    }
  ]
}
```

Loop TDD para IA:

1. Escribir/editar un `Public Function Test_*() As String`.
2. Agregar el caso a `tests.vba.json`.
3. Importar módulos cambiados:
   ```powershell
   node cli.js import Test_Clientes Clientes --access "CONDOR.accdb"
   ```
4. Ejecutar:
   ```powershell
   node cli.js test-vba --access "CONDOR.accdb" --json
   ```
5. Si `phase=compile`, corregir compilación. No discutas assertions todavía.
6. Si `phase=tests`, leer `results[*].failures`, `run.payload`, `run.logs` y corregir comportamiento.
7. Reimportar y repetir hasta `ok=true`.

### Ejemplo completo mínimo

Módulo de producción `src/modules/Calculadora.bas`:

```vb
Option Compare Database
Option Explicit

Public Function Sumar(ByVal a As Long, ByVal b As Long) As Long
    Sumar = a + b
End Function
```

Módulo de test `src/modules/Test_Calculadora.bas`:

```vb
Option Compare Database
Option Explicit

Public Function Test_SumarBasico() As String
    On Error GoTo EH

    Dim resultado As Long
    resultado = Sumar(2, 3)

    If resultado = 5 Then
        Test_SumarBasico = "{""ok"":true,""value"":5,""logs"":[""2+3=5""]}"
    Else
        Test_SumarBasico = "{""ok"":false,""error"":""resultado esperado 5"",""value"":" & CStr(resultado) & "}"
    End If
    Exit Function

EH:
    Test_SumarBasico = "{""ok"":false,""error"":""" & Replace(Err.Description, """", "'") & """}"
End Function
```

`tests.vba.json`:

```json
{
  "tests": [
    {
      "name": "suma básica",
      "procedure": "Test_SumarBasico",
      "expect": { "ok": true, "value": 5 },
      "tags": ["calculadora", "unit"]
    }
  ]
}
```

Comandos:

```powershell
node cli.js import Calculadora Test_Calculadora --access "MiProyecto.accdb"
node cli.js test-vba --access "MiProyecto.accdb" --json
```

### Cómo debe razonar la IA con el reporte

Si recibe:

```json
{ "ok": false, "phase": "compile", "compile": { "component": "Clientes", "line": 18 } }
```

Debe corregir compilación en `Clientes` y repetir. Esto suele ser firma incorrecta, argumento faltante, variable no declarada, referencia rota o sintaxis.

Si recibe:

```json
{ "ok": false, "phase": "tests", "failed": 1 }
```

Debe mirar:

- `results[*].failures` para saber qué expectation falló.
- `results[*].run.payload` para el valor real.
- `results[*].run.logs` para trazas devueltas por VBA.

### Límites conocidos

- Los argumentos COM deben ser simples: `string`, `number`, `boolean`, `null`.
- Los argumentos COM están limitados a 10 por llamada. Si necesitás más, empaquetalos en un JSON String.
- Para objetos complejos, pasá un `String` JSON como argumento y parsealo en VBA.
- El compile gate depende de Access/VBE instalado y con permisos de automatización.
- El diagnóstico de línea de compilación es best-effort: Access puede no exponer siempre ubicación exacta.
- Si una función muestra `MsgBox`, abre formularios modales o espera input, el runner puede bloquearse. Los tests deben evitar UI.

---

## Estructura de archivos exportados

```
<DestinationRoot>/
├── modules/          # Tipo 1 (vbext_ct_StdModule)   → .bas
├── classes/          # Tipo 2 (vbext_ct_ClassModule)  → .cls
└── forms/            # Tipo 3/100 (formularios)
    ├── Form_X.form.txt   # UI + código (SaveAsText, acForm=2)
    └── Form_X.cls        # Solo código VBA
```

Ejemplo real:
```
src/modules/VariablesGlobales.bas
src/classes/CUsuario.cls
src/forms/Form_FormGestion.form.txt
src/forms/Form_FormGestion.cls
```

⚠️ `export` y `export-all` **no escriben en la raíz pelada** del destino.  
Siempre exportan dentro de subcarpetas tipadas:

- `classes/`
- `modules/`
- `forms/`

Ejemplo:

```powershell
node cli.js export DocumentoServicio --access "CONDOR.accdb" --destination "C:\temp\prueba"
```

escribirá en:

```text
C:\temp\prueba\classes\DocumentoServicio.cls
```

---

## Comportamiento de formularios

Los formularios Access tienen tratamiento especial respecto a módulos y clases:

- **Export**: `Application.SaveAsText(2, nombreSinPrefixForm_, ruta)`. El nombre del objeto Access es el nombre del VBComponent **sin** el prefijo `Form_`.
- **Import**: `Application.LoadFromText(2, nombreSinPrefixForm_, ruta)`. Nunca usa `VBComponents.Import()`.
- **Fallback en export**: si `SaveAsText` falla, usa `component.Export()` y registra el aviso.
- Los formularios también generan un `.cls` paralelo con solo el código VBA para facilitar diff.
- **Búsqueda automática**: si se pasa `frmNombre` (sin prefijo `Form_`), el sistema busca automáticamente `Form_frmNombre.form.txt`. funciona en ambos sentidos: se puede usar `frmSplash` o `Form_frmSplash`.

---

## Regla de oro: `import` canónico, código en .cls, UI en .form.txt

**Nunca editar el CodeBehind del `.form.txt` directamente.** El flujo correcto es:

1. **CAMBIO DE CÓDIGO VBA** → editar **SOLO el `.cls`**
2. **CAMBIO DE UI** (propiedades de controles, layout) → editar **SOLO el `.form.txt`**
3. **Antes de importar** (modo Auto/canónico) → si existe `.form.txt`, el handler sincroniza automáticamente el CodeBehind del `.form.txt` con el contenido del `.cls`

### Regla dura de seguridad para `import-code`

Si el `.cls` **parece code-behind de formulario/reporte** pero la skill **no puede resolver un document module existente** dentro de Access, **debe abortar**.

No está permitido hacer fallback a `VBComponents.Import()` en ese caso, porque eso contamina el binario creando módulos espurios como:
- `Módulo1`
- `Módulo2`

Si pasa, la acción correcta es:
- verificar el nombre real del formulario/document module
- exportar primero el objeto correcto
- o usar `import` / `import-form` según corresponda

### Por qué

El `.cls` es el archivo "maestro" para código VBA. El `.form.txt` tiene dos secciones:
- **UI** (antes de `CodeBehind`): propiedades de controles, layout
- **CodeBehind** (después de `CodeBehind`): código VBA — **es un espejo del `.cls`**

Cuando Access exporta un formulario, guarda el código en ambas secciones por separado. Si la IA modifica solo el `.cls` pero el CodeBehind del `.form.txt` está desincronizado, los cambios no se aplican correctamente.

### Cómo funciona el sync automático

En `import-modules` (handler.js), antes de invocar VBAManager.ps1 con `-ImportMode Code` o `-ImportMode Auto`:
1. Se lee el contenido del `.cls`
2. Se reemplaza la sección `CodeBehind` del `.form.txt` con ese contenido
3. Se importa el `.form.txt` ya sincronizado a Access

### Verificación manual

Para verificar que `.cls` y `CodeBehind` coinciden sin hacer import:

```powershell
# Un formulario concreto
node cli.js verify-code Form_subfrmDatosCDCA_Generales

# Varios a la vez
node cli.js verify-code Form_frmGestion Form_subfrmDetalle

# Todos los formularios/reportes con .cls
node cli.js verify-code
```

Salida:
- `✅ Form_X: en sync` — el `.cls` y el CodeBehind del `.form.txt` son idénticos
- `❌ Form_X: DESINCRONIZADO` seguido de un diff unificado con el contenido de ambos

El comando termina con exit code 1 si detecta alguna desincronización, lo que permite usarlo en scripts o pipelines de validación.

---

## Flujo de modificación de UI de formularios por la IA

Este es un caso de uso de primera clase del skill: **la IA modifica la interfaz de un formulario editando su `.form.txt` y el cambio se aplica a la BD**.

### El ciclo completo

```
export Form_X  →  IA edita src/forms/Form_X.form.txt  →  import Form_X
```

El `.form.txt` contiene tanto la definición de controles y propiedades (UI) como el código VBA. Al importarlo con `LoadFromText`, Access recarga el formulario completo — UI y código — reflejando todos los cambios.

### Comandos para este flujo

```powershell
# 1. Exportar el formulario a src/ (si no está ya exportado)
node cli.js export Form_FormGestion

# 2. La IA edita src/forms/Form_FormGestion.form.txt
#    (propiedades de controles, nuevos botones, layout, etc.)

# 3. Importar el formulario modificado a la BD
node cli.js import Form_FormGestion

# Varios formularios a la vez:
node cli.js import Form_FormGestion Form_subfrmDetalle Form_FormBusqueda
```

### En modo watch (automático)

Con `watch` activo, cada vez que la IA guarda un `.form.txt` se importa automáticamente a la BD sin ninguna intervención:

```powershell
node cli.js watch
# A partir de aquí: la IA edita y guarda .form.txt → se importa solo
```

### Regla crítica: nunca crear un `.form.txt` desde cero

La IA **debe partir siempre de un `.form.txt` exportado con `export`**, nunca generarlo de cero. El formato es propietario de Access — si se introduce una línea mal formada, `LoadFromText` falla. El flujo correcto es siempre exportar primero para tener la base correcta y luego modificar.

### Qué puede modificar la IA en el `.form.txt`

El archivo tiene esta estructura general:

```
Version =21
VersionRequired =20
Begin Form
    Width = ...
    Caption = "..."
    Begin Label etiqueta1
        Left = ...
        Top = ...
        Caption = "Nombre:"
    End
    Begin TextBox txtNombre
        Left = ...
        Width = ...
        ControlSource = "NombreCampo"
    End
    ...
End
CodeBehind
    ' Aquí va el código VBA del formulario
```

La IA puede modificar con seguridad:
- Propiedades de controles existentes (`Left`, `Top`, `Width`, `Height`, `Caption`, `Visible`, etc.)
- Valores de `ControlSource`, `RowSource`, `DefaultValue`
- Añadir nuevos controles copiando la estructura de uno existente

Debe evitar:
- Editar manualmente la sección `CodeBehind` del `.form.txt` (el skill la sincroniza desde el `.cls`)
- Cambiar el `GUID` o `Checksum` del formulario (Access los regenera en el siguiente export)
- Renombrar controles que tengan referencias en el código VBA sin actualizar también el código
- Alterar la estructura de bloques `Begin`/`End` de forma que queden mal anidados

---

## Comportamiento de sesión

El skill persiste el estado en `.access-vba-skill/session.json` en la raíz del proyecto:

```json
{
  "active": true,
  "startedAt": "2025-01-15T09:00:00.000Z",
  "accessPath": "C:\\proyecto\\MiBD.accdb",
  "destinationRoot": "C:\\proyecto\\src",
  "modulesPath": "C:\\proyecto\\src",
  "changedModules": ["Utilidades", "Form_FormGestion"],
  "lastSyncAt": "2025-01-15T09:45:00.000Z",
  "pendingModules": [],
  "watcherPid": null
}
```

Esto permite que `import`, `sync`, `end` y `status` funcionen sin mantener un proceso largo vivo.

---

## Gestión de Access (headless)

Antes de abrir la BD el PS1 deshabilita temporalmente mediante DAO:
- **AllowBypassKey**: habilitado para acceso sin restricciones
- **StartupForm**: eliminado para que no abra el formulario de inicio
- **AutoExec**: renombrado a `AutoExec_TraeBackup` para evitar ejecución

La BD se abre con:
```powershell
$access.Visible = $false
$access.UserControl = $false
$access.AutomationSecurity = 1   # desactiva seguridad de macros
$access.DoCmd.SetWarnings($false) # suprime diálogos de Access
```

Al cerrar, todos los valores se restauran al estado original.

---

## Estructura del paquete

```
<projectRoot>/
  access-vba-sync/
    VBAManager.ps1     # Motor PowerShell (Export/Import/Fix-Encoding/Generate-ERD)
    handler.js         # Lógica Node.js (sesión, debounce, watcher, métodos del skill)
    cli.js             # Interfaz de comandos
    README.md          # Documentación de uso
    SKILL.md           # Este documento (especificación)
```

> El skill vive en su carpeta pero opera con `projectRoot = cwd`, de forma que `src/` y `docs/ERD/` quedan en el proyecto y no dentro del skill.

---

## Flujo de trabajo esperado

### Nueva feature/fix
```
start → (generate-erd) → IA edita src/ → watch|import → compilar VBE → end
```

1. `start` — Export total, snapshot base
2. `generate-erd` — Contexto de datos para la IA (opcional)
3. IA modifica archivos en `src/`
4. `watch` (automático) o `import <módulos>` (manual)
5. Compilar: **Abre Access → VBE → Debug → Compile**
6. `end` — Sync final + export final opcional + resumen

---

## Casos límite cubiertos

- Varias BDs en CWD → elección determinista (alfabético) + warning
- Ruta relativa en `--access` → se resuelve contra CWD
- Nombre de módulo sin extensión → `Resolve-ImportFileForModule` busca por prioridad: `.form.txt` > `.frm` > `.cls` > `.bas`
- Formularios sin prefijo `Form_` → detectados por tipo de componente (tipo 3 o 100)
- Guardados masivos simultáneos → batching con debounce configurable
- Access abierto/bloqueado → error claro, no loops infinitos
- BD con contraseña → `--password` propagado a todas las operaciones DAO y COM

---

## Pruebas mínimas

- `start` con BD única → crea `src/modules/`, `src/classes/`, `src/forms/`
- `start` sin BD en CWD → error claro
- `export Form_X` → genera `src/forms/Form_X.form.txt` y `src/forms/Form_X.cls`
- `import Form_X` → reimporta el formulario sin diálogos
- `watch`: editar un `.bas` → se importa automáticamente tras debounce
- `watch`: editar un `.form.txt` → se importa automáticamente
- `fix-encoding --location Src` → solo toca archivos en `src/`
- `import-all` → importa todos los archivos de `src/` sin intervención
- `generate-erd` → genera `docs/ERD/NombreBackend.md`
- `end` → export final + resumen con módulos tocados
