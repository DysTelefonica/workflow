---
name: access-vba-sync
description: >
  Navaja suiza global para workflows de Microsoft Access/VBA con código fuera del binario,
  sincronización bidireccional con archivos en disco y soporte para Git.
  Trigger: úsala SIEMPRE que una IA vaya a modificar, haya modificado o deba sincronizar
  código VBA/Access en proyectos que usen este workflow, especialmente antes de importar
  cambios a la BD o al preparar snapshots/exportaciones.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "2.0"
---

# access-vba-sync

## Cuándo usar esta skill

Usá esta skill en cualquier proyecto Access/VBA donde el código viva temporalmente fuera del binario para poder:

- trabajar con Git
- editar módulos con IA
- exportar/importar entre la BD Access y archivos en disco
- generar contexto técnico (ERD)
- crear un sandbox local de backends vinculados

### Trigger principal

**Si una IA cambió cualquier archivo VBA exportado (`.bas`, `.cls`, `.form.txt`, `.report.txt`, `.frm`) o va a cambiarlo, esta skill DEBE usarse.**

Porque al final esos cambios tienen que volver al binario de Access mediante import/sync.

---

## Qué problema resuelve

Access guarda formularios, módulos y clases dentro de un binario (`.accdb`, `.mdb`, etc.). Eso dificulta:

- versionado real con Git
- diffs útiles
- edición asistida por IA
- revisión de cambios
- documentación del modelo de datos

Esta skill convierte ese workflow en uno repetible:

1. exportás desde Access a disco
2. editás archivos fuera del binario
3. sincronizás/importás de vuelta a Access
4. compilás manualmente en VBE
5. opcionalmente generás ERD o armás sandbox local

---

## Regla de oro

### El código fuente de trabajo está en disco; la verdad ejecutable final está en Access.

Eso significa:

- la IA **edita archivos exportados**
- luego hay que **importarlos** a la BD
- el ciclo no termina hasta que los cambios estén dentro de Access

---

## Regla crítica sobre `start` y `export-all`

### EXTREMO CUIDADO

`start` y `export-all` vuelcan el código del binario Access hacia disco.

Si la IA ya modificó archivos exportados y todavía no los importaste, ejecutar:

- `start`
- `export-all`
- o `export` sobre esos módulos

puede **pisar en disco el trabajo nuevo con código viejo del binario**.

## Interpretación obligatoria para IA

### Antes de ejecutar `start`, `export-all` o `export <módulos>` preguntate:

**¿Estoy seguro de que no voy a sobrescribir cambios locales más nuevos que todavía no fueron importados?**

Si la respuesta no es un **sí claro**, no exportes.

### Regla operativa

- **Si la IA acaba de modificar archivos en `src/`** → normalmente corresponde `import`, no `start` ni `export-all`
- **Si querés tomar snapshot inicial desde la BD** y todavía no hubo cambios locales → `start` o `export-all` sí
- **Si hay duda entre exportar o importar** → priorizá proteger el trabajo local y no sobrescribirlo

---

## Arquitectura real del workflow

Esta skill tiene 3 capas. No las confundas.

### 1. `VBAManager.ps1`
Motor PowerShell que opera sobre Access/DAO/COM.

Acciones reales:

- `Export`
- `Import`
- `Fix-Encoding`
- `Generate-ERD`
- `Sandbox`

### 2. `handler.js`
Orquesta sesión, debounce, watcher, sync de CodeBehind y UX del workflow.

### 3. `cli.js`
Interfaz de comandos para usar el workflow desde terminal.

---

## Archivos clave

En un proyecto típico, esta skill vive globalmente pero opera sobre el proyecto actual (`cwd`).

### Archivos del skill
- `skills/access-vba-sync/VBAManager.ps1`
- `skills/access-vba-sync/handler.js`
- `skills/access-vba-sync/cli.js`
- `skills/access-vba-sync/SKILL.md`

### Archivos del proyecto afectados
- `<projectRoot>/*.accdb|*.accde|*.mdb|*.mde`
- `<projectRoot>/src/modules/*.bas`
- `<projectRoot>/src/classes/*.cls`
- `<projectRoot>/src/forms/*.form.txt`
- `<projectRoot>/src/forms/*.cls`
- `<projectRoot>/src/reports/*.report.txt`
- `<projectRoot>/src/reports/*.cls`
- `<projectRoot>/docs/ERD/*.md`
- `<projectRoot>/.access-vba-skill/session.json`

---

## Capacidades reales

## Export
Exporta desde Access a disco.

### Qué exporta
- módulos estándar → `.bas`
- clases → `.cls`
- formularios → `.form.txt` + `.cls`
- reportes → `.report.txt` + `.cls`

### Cuándo usarlo
- snapshot inicial de una BD todavía no exportada
- snapshot final explícito y consciente
- export selectivo de un módulo que todavía no está en disco

### Cuándo NO usarlo
- después de que la IA ya editó archivos locales no importados
- como reflejo automático “por las dudas”

---

## Import
Importa cambios desde disco hacia la BD Access.

### Este es el comando normal después de cambios hechos por IA.

Usalo cuando:
- la IA modificó `.bas`
- la IA modificó `.cls`
- la IA modificó `.form.txt`
- la IA modificó `.report.txt`
- querés sincronizar de vuelta al binario

### Modos
- `Auto`
- `Form`
- `Code`

---

## Fix-Encoding
Corrige problemas de encoding:
- BOM en disco
- resync hacia Access

No lo uses como rutina por defecto si no hay problema real de encoding.

---

## Generate-ERD
Genera documentación Markdown de tablas, campos, índices, PKs y relaciones.

Usalo cuando la IA necesite contexto del modelo de datos.

---

## Sandbox
Copia backends vinculados al lado del frontend y revincula a copias locales.

Usalo cuando querés trabajar aislado de producción o de una red compartida.

### Importante
`Sandbox` es una acción real del PS1. No es solo una capa del CLI.

---

## Comandos CLI que la IA debe conocer

| Comando | Uso principal | Riesgo |
|---|---|---|
| `start` | export inicial + sesión | **ALTO**: puede sobrescribir cambios locales |
| `watch` | auto-import al guardar archivos | medio |
| `export <Mod...>` | export selectivo | **ALTO** si hay cambios locales no importados |
| `export-all` | export total | **ALTO** si hay cambios locales no importados |
| `import <Mod...>` | importar módulos selectivos | bajo |
| `import-form <Mod...>` | importar UI de documentos Access (`.form.txt` / `.report.txt`) | bajo |
| `import-code <Mod...>` | importar solo código | bajo |
| `import-all` | importar todo `src/` | medio |
| `fix-encoding [Mod...]` | corregir encoding | medio |
| `generate-erd` | documentación técnica | bajo |
| `sandbox` | sandbox local de backends | medio |
| `list-objects` | inspección del frontend real | bajo |
| `exists <Mod>` | verificar si un objeto/módulo existe de verdad | bajo |
| `status` | inspeccionar sesión | bajo |
| `end` | cierre de sesión + export final opcional | medio |

---

## Decisión rápida: qué comando usar

## Caso A — La IA ya cambió archivos en `src/`
### Usar:
- `import <módulos>`
- `import-code <módulos>`
- `import-form <módulos>`
- `import-all` si está controlado

### No usar por defecto:
- `start`
- `export-all`
- `export`

---

## Caso B — Recién arrancás y querés sacar el código fuera del binario
### Usar:
- `start`
- o `export-all`

### Condición obligatoria:
- no debe haber trabajo local más nuevo que el binario

---

## Caso C — Necesitás contexto de datos para la IA
### Usar:
- `generate-erd`

---

## Caso D — Querés trabajar sin tocar backends reales
### Usar:
- `sandbox`

## Caso E — La IA no sabe si un formulario/reporte/módulo existe realmente en la BD
### Usar:
- `list-objects`
- `exists <Modulo>`

### Regla operativa
Antes de intentar crear o importar un subform/reporte dudoso, inspeccioná el frontend real.  
No adivines si el objeto existe.

---

## Formularios: reglas críticas

Los formularios tienen dos artefactos principales:

- `.form.txt` → UI + definición completa del objeto Access
- `.cls` → código VBA del formulario

Los reportes siguen la misma idea:

- `.report.txt` → UI + definición completa del reporte
- `.cls` → código VBA del reporte

## Regla obligatoria

### Código VBA del formulario
Editar **preferentemente el `.cls`**.

### UI / layout / controles / propiedades
Editar **el `.form.txt`**.

---

## Sobre `CodeBehind`

El handler sincroniza automáticamente el `CodeBehind` del `.form.txt` o `.report.txt` a partir del `.cls` antes de ciertos imports.

### Consecuencia práctica
- para cambios de código, el archivo maestro debe ser el `.cls`
- para cambios visuales, el archivo maestro es el `.form.txt` o `.report.txt`

### No hagas esto
- no generes un `.form.txt` desde cero
- no generes un `.report.txt` desde cero
- no inventes sintaxis del formato Access
- no asumas que editar a mano el `CodeBehind` dentro del `.form.txt` es el flujo ideal

### Sí hacé esto
- partí siempre de un `.form.txt` exportado por Access
- partí siempre de un `.report.txt` exportado por Access si trabajás con reportes
- modificá solo lo necesario
- si el cambio es de lógica, preferí el `.cls`

---

## Forms vs Reports

### Regla importante para IA
No asumas que todo document module es un formulario.

Access distingue al menos:
- forms
- reports

Si el motor necesita importar/exportar objetos de documento, debe resolver correctamente el tipo Access real. No reforzar el supuesto simplista de que todo `vbext_ct_Document` = form.

### Consecuencia práctica
- `import-form` también puede terminar importando reportes si el archivo fuente real es `.report.txt`
- `import-code` sincroniza `CodeBehind` tanto para forms como para reports antes de importar

---

## Compilación

Después de importar cambios, la skill/CLI puede **recordar** al usuario que compile.

### Pero la compilación es manual

Abrir Access → VBE → `Debug -> Compile`

La skill no sustituye esa validación.

---

## Gestión de sesión

La sesión se persiste en:

- `.access-vba-skill/session.json`

Eso permite:
- `watch`
- `status`
- `end`
- tracking de módulos cambiados y pendientes

La IA puede usar `status` para entender el estado antes de decidir acciones destructivas.

## Introspección del frontend

La skill ya puede inspeccionar la BD real para que la IA no trabaje a ciegas.

### `list-objects`
Lista:
- forms
- reports
- modules
- classes
- documentModules

Uso:

```bash
node cli.js list-objects --access MiBD.accdb
node cli.js list-objects --access MiBD.accdb --json
```

### `exists <Modulo>`
Responde si un nombre dado existe en la BD y sugiere el tipo de import.

Uso:

```bash
node cli.js exists Form_subfrmDatosPCSUB_Propuesta --access MiBD.accdb
node cli.js exists Form_subfrmDatosPCSUB_Propuesta --access MiBD.accdb --json
```

### Cuándo usarlo
- cuando un `import-form` falla y no sabés si el objeto existe
- cuando un `import-code` depende de un objeto base en Access
- cuando la IA necesita distinguir entre:
  - objeto Access existente
  - VBComponent existente
  - módulo/clase existente

---

## Reglas de seguridad operacional

## 1. No exportar a ciegas
Nunca uses `start` o `export-all` si podés pisar trabajo local más nuevo.

## 2. No asumir que la BD está cerrada
Idealmente no debería estar abierta por el usuario, pero el PS1 intenta cerrar la instancia objetivo si la detecta.

## 3. No asumir soporte mágico
Si algo depende de:
- Access instalado
- COM
- DAO
- rutas válidas
- contraseñas

verificalo.

## 4. No prometer fallbacks que no existen
Si `SaveAsText` o `LoadFromText` fallan, no supongas que el script siempre tiene fallback silencioso.

## 5. No inventar comandos
Si un comando no existe en `cli.js`, no lo uses.

## 6. No asumir semántica inexistente en `--keep_sidecars`
Hoy `--keep_sidecars` es esencialmente informativo: los sidecars del sandbox se conservan igualmente. No bases lógica crítica en que ese flag limpie o no limpie archivos.

---

## Anti-patrones

### No hagas esto
- ejecutar `start` después de que la IA ya cambió archivos en `src/`
- usar `export-all` para “sincronizar” cambios locales al binario
- crear `.form.txt` desde cero
- crear `.report.txt` desde cero
- asumir que `Sandbox` limpia sidecars automáticamente si no está implementado así
- asumir que reportes siguen exactamente el mismo flujo que formularios
- mezclar en la explicación lo que hace el PS1 con lo que hace el handler sin aclararlo

---

## Workflow recomendado para IA

## Flujo normal tras cambios hechos por IA

```bash
node cli.js status
node cli.js import <Mod1> <Mod2>
# o import-code / import-form según el caso
```

Luego:
- pedir o recordar compilación manual en Access/VBE

---

## Flujo de inicio seguro

```bash
node cli.js status
# verificar que no hay cambios locales pendientes de proteger
node cli.js start --access <BD>
```

---

## Flujo con formularios

### Cambio de código
```bash
node cli.js import-code Form_MiFormulario
```

### Cambio de UI/layout
```bash
node cli.js import-form Form_MiFormulario
```

### Cambio de UI/layout en reporte
```bash
node cli.js import-form Report_MiReporte
```

### Si no estás seguro
```bash
node cli.js import Form_MiFormulario
```

---

## Flujo con ERD

```bash
node cli.js generate-erd --backend <ruta_backend>
```

---

## Flujo con sandbox

```bash
node cli.js sandbox --access <frontend.accdb>
```

---

## Checklist mental antes de actuar

Antes de ejecutar cualquier comando, la IA debe chequear:

- [ ] ¿Estoy exportando o importando?
- [ ] ¿Hay riesgo de sobrescribir cambios locales nuevos?
- [ ] ¿El cambio está en `.cls`, `.form.txt`, o ambos?
- [ ] ¿Necesito `import`, `import-code` o `import-form`?
- [ ] ¿Necesito contexto de datos (`generate-erd`)?
- [ ] ¿Estoy trabajando contra producción o necesito `sandbox`?

---

## Comandos canónicos

```bash
# Export inicial (PELIGROSO si hay cambios locales no importados)
node cli.js start --access MiBD.accdb

# Export total consciente (PELIGROSO si hay cambios locales no importados)
node cli.js export-all --access MiBD.accdb

# Import selectivo tras cambios hechos por IA
node cli.js import Mod_Utilidades Form_MiFormulario

# Import solo código
node cli.js import-code Form_MiFormulario

# Import solo UI/formulario
node cli.js import-form Form_MiFormulario

# Auto-sync mientras se edita
node cli.js watch --access MiBD.accdb

# Generar ERD
node cli.js generate-erd --backend MiBD_Datos.accdb

# Crear sandbox
node cli.js sandbox --access MiBD.accdb --backend_password xxx

# Ver estado
node cli.js status

# Cerrar sesión
node cli.js end
```

---

## Lo que la IA debe recordar siempre

### Esta skill es la navaja suiza del workflow Access/VBA.

No es un detalle accesorio del proyecto.
En estos proyectos, si la IA modifica código exportado, **esta skill forma parte obligatoria del cierre correcto del trabajo**.

Si la IA tocó código y no lo importó a Access, el trabajo está incompleto.

Y si la IA ejecuta `start` o `export-all` sin pensar, puede destruir su propio trabajo local pisándolo con el código viejo del binario.

Ese es el error que NO se puede cometer.
