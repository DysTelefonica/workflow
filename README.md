# Skills

Colección de skills para OpenCode/Trae Agent. Cada carpeta es una skill independiente.

## Skills disponibles

| Skill | Descripción |
|---|---|
| `access-vba-sync` | Sincronización bidireccional VBA entre Access y sistema de archivos |
| `access-sandbox` | Provisiona backends Access locales y reescribe tablas vinculadas |
| `access-query` | Ejecuta SQL contra backends Access (.accdb) de proyectos VBA |
| `pdf` | Lectura, fusión, división y OCR de PDFs |
| `obsidian-*` | Gestión de bóvedas, notas, bases y canvas en Obsidian |
| `skill-creator` | Framework para crear y validar nuevas skills para agentes IA |
| `jira-kanban-portfolio` | Convenciones Jira para Kanban por espacios/proyectos reales y PORTFOLIO transversal |
| `jira-kanban-ticketing` | Redacción operable de tickets Jira para Kanban con criterios de aceptación y validación |

## Estructura de una skill

```
skill-name/
    SKILL.md         # Descripcion y objetivos
    references/      # Documentacion adicional
    scripts/         # Scripts de automatizacion
```

## Uso rapido

```powershell
# Sync backends de produccion a local
.\access-sandbox\sync.bat
```
