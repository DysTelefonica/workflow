# SKILL.md — Skill para workflow Access/VBA (Export → Trabajo → Sync → Compilar → Cierre)

## Objetivo
Definir un **skill** (implementación a realizar por otra IA) que automatice el workflow de desarrollo en un proyecto Microsoft Access/VBA:

1) **Al inicio** de una nueva feature/fix: **Exportar TODOS los módulos** del proyecto VBA a disco (snapshot base).
2) Se trabaja sobre la mejora editando los archivos exportados (normalmente con IA).
3) **Todo módulo modificado por la IA debe sincronizarse** (Import) hacia el VBA real de la BD.
4) Tras cada sincronización, el skill **debe proponer al usuario compilar** el proyecto en el VBE.
5) **Al cerrar** la tarea (fin de sesión): export final opcional (snapshot consistente) + resumen.

> El skill debe ser **autocontenido**: incluir dentro **VBAManager.ps1** y todo lo necesario para ejecutarse.

---

## Alcance y supuestos
- Entorno: **Windows** con Microsoft Access instalado (automatización COM).
- El repositorio contiene una BD Access (`.accdb/.accde/.mdb/.mde`) en la raíz del proyecto, o el usuario la pasa por parámetro.
- La exportación se guarda bajo una carpeta configurable `src/`.
- Se asume que `VBAManager.ps1` soporta:
  - `-Action Export|Import|Fix-Encoding`
  - `-AccessPath <ruta>`
  - `-DestinationRoot <carpeta>`
  - `-ModuleName <string[]>` (múltiples).
  Si NO soporta array, el skill debe iterar e invocar Import 1×módulo.

---

## Requisitos funcionales (MUST)
### R1. Inicio de sesión (start)
- Detectar `AccessPath`:
  - Si el usuario lo pasa: usarlo (aceptar rutas relativas a project root).
  - Si no: autodetectar en project root: `.accdb/.accde/.mdb/.mde`.
    Si hay varias, elegir determinista (alfabético) y avisar.
- Ejecutar: `VBAManager.ps1 -Action Export -AccessPath ... -DestinationRoot ...`
- Persistir estado de sesión en disco (para que `sync/end/status` funcionen sin mantener proceso vivo):
  - accessPath, destinationRoot, modulesPath, startedAt, changedModules.

### R2. Ruta real de módulos exportados
El skill exporta e importa directamente en la carpeta destino:

`<DestinationRoot>/*.bas|*.cls|*.frm`

Ejemplo:
`src/Utilidades.bas`

### R3. Sincronización (sync/import)
- Dado un conjunto de módulos (por nombre), ejecutar Import **solo de esos**:
  - `VBAManager.ps1 -Action Import -AccessPath ... -DestinationRoot ... -ModuleName A B C`
- Registrar en el estado: `changedModules += módulos`.
- Tras importar: **mostrar instrucción explícita** al usuario:
  - “Abre Access → VBE → Debug → Compile”.

### R4. Auto-sync durante el trabajo (watch)
- Vigilar `modulesPath` (que coincide con `DestinationRoot`) y detectar cambios en:
  - `.bas`, `.cls` (y opcional `.frm` si tu proyecto lo usa).
- Al cambiar un archivo:
  - Derivar `ModuleName` = basename sin extensión.
  - Hacer debounce/batching (ej. 500–1000 ms) y luego Import de todos los módulos tocados en esa ventana.
- En `unlink` (borrado): avisar (no se puede borrar módulo en VBA automáticamente de forma segura).

### R5. Fin de sesión (end)
- Parar watcher si está activo.
- Si hay cambios pendientes: hacer sync final.
- Export final opcional (configurable): `-Action Export`.
- Imprimir resumen: nº módulos sincronizados + lista.

### R6. Comandos mínimos del skill
El skill debe exponer al menos:
- `start` (export inicial + estado)
- `watch` (start si no hay sesión + auto-sync)
- `sync`/`import <Mod...>` (import manual por lista)
- `end` (cierre + export final opcional)
- `status` (estado de sesión)

---

## Requisitos no funcionales (SHOULD)
- No bloquear el hilo principal: ejecutar PowerShell como proceso hijo (capturar stdout/stderr).
- Log claro y accionable (qué módulo se importó y por qué).
- Fallos: si Import falla, mostrar el error + stdout/stderr del PS1.
- Configurable por fichero (ej. `skill.config.json`) o flags:
  - destinationRoot (default `src`)
  - debounceMs
  - autoExportOnStart / autoExportOnEnd
- No depender de servicios externos; todo local.

---

## Estructura propuesta del paquete del skill
<projectRoot>/
access-vba-sync/
SKILL.md
scripts/
  VBAManager.ps1
  handler.js
  cli.js
  package.json

> Importante: el skill vive en su carpeta, pero se ejecuta con `projectRoot = cwd` (la raíz del repo), para que `src/` quede en el proyecto y no dentro del skill.

---

## Flujo de trabajo esperado (integración)
### Nueva feature/fix
1) `start` → Export total a `src/`
2) IA modifica archivos en esa carpeta.
3) `watch` (o `sync` al terminar) → Import de módulos modificados.
4) Usuario compila en VBE cuando el skill lo recuerde.
5) `end` → sync final + export final opcional.

---

## Casos límite que el skill debe cubrir
- Varias BDs en root → elección determinista + warning.
- Ruta relativa de AccessPath (como el resto de comandos del proyecto).
- Módulos con mismo nombre en diferentes extensiones (preferir el archivo cambiado; importar por nombre).
- Cambios masivos (muchos guardados) → batching.
- Access abierto/bloqueado → error claro (no loops infinitos).

---

## Pruebas mínimas
- Start con BD única y sin BD.
- Export crea `src/*.bas|*.cls|*.frm`.
- Watch: editar un `.bas` y confirmar Import.
- Import manual con 2 módulos (array).
- End: export final + resumen.
