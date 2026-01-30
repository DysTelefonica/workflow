# SKILL.md — Skill Corrector UTF-8 para VBA

## Objetivo
Herramienta especializada para limpiar y corregir problemas de codificación en archivos de código fuente VBA exportados desde Microsoft Access. Elimina BOMs (Byte Order Marks) y estandariza a UTF-8 para evitar caracteres extraños (mojibake) y facilitar el control de versiones.

## Problema que resuelve
- **Encoding Incorrecto**: Access exporta en Windows-1252 o UTF-16LE, lo que causa problemas en Git y editores modernos.
- **Caracteres BOM**: Elimina los bytes de marca de orden (BOM) que ensucian los diffs y causan errores en algunos intérpretes.
- **Mojibake**: Corrige caracteres mal interpretados (tildes, eñes) detectando la codificación original.

## Herramientas Incluidas
### `scripts/clean_vba.py`
Script Python ligero y eficiente para limpieza automática.

**Uso:**
```bash
python utf8-corrector/scripts/clean_vba.py <archivo_o_directorio> [archivo2 ...]
```

**Opciones:**
- `--dry-run`: Simula la conversión sin modificar los archivos.

**Ejemplos:**
```bash
# Limpiar un archivo específico
python utf8-corrector/scripts/clean_vba.py src/Modulo1.bas

# Limpiar todo un directorio (busca .bas, .cls, .frm)
python utf8-corrector/scripts/clean_vba.py src/
```

## Requisitos
- Python 3.x (solo librería estándar)

## Integración
Este skill está diseñado para ser invocado por agentes de IA o pipelines de CI/CD antes de realizar commits de código VBA.
