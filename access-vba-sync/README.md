# skill_access_vba_sync

Skill local autocontenido para tu workflow Access/VBA:

1) `start` → Export inicial (snapshot)  
2) Editas archivos exportados en `src/`  
3) `watch` auto-importa cada guardado (`.bas/.cls` y opcional `.frm`)  
4) Tras cada import: te recuerda “Abre Access → VBE → Debug → Compile”  
5) `end` → sync final + export final opcional + resumen  

## Requisitos

- Windows
- Microsoft Access instalado (automatización COM)
- PowerShell
- Node.js 18+

## Instalación

Desde la raíz del proyecto (CWD = projectRoot):

```powershell
cd skill_access_vba_sync
npm install
```

## Uso

Siempre ejecuta los comandos desde la raíz del proyecto (donde está la `.accdb`):

### Start (export inicial)
```powershell
node skill_access_vba_sync/cli.js start --access "MiBD.accdb"
```

Si no pasas `--access`, autodetecta en CWD (`.accdb/.accde/.mdb/.mde`). Si hay varias, elige la primera por orden alfabético y avisa.

### Watch (lo normal)
```powershell
node skill_access_vba_sync/cli.js watch --access "MiBD.accdb" --debounce_ms 800
```

Edita módulos en:
`src/*.bas|*.cls|*.frm`

### Import / Sync manual por módulos
```powershell
node skill_access_vba_sync/cli.js import Utilidades Validaciones
```

### Status
```powershell
node skill_access_vba_sync/cli.js status
```

### End
```powershell
node skill_access_vba_sync/cli.js end
```

## Flags útiles

- `--destination_root src` (default: `src`)
- `--auto_export_on_end false`
