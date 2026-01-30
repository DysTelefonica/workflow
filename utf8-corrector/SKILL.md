# SKILL.md — Skill Corrector UTF-8

## Objetivo
Proporcionar una herramienta capaz de detectar y corregir problemas de codificación en archivos de texto, asegurando que todo el contenido esté en UTF-8 válido.

## Funcionalidades Principales
1. **Detección**: Identificar archivos con codificación incorrecta (ANSI, Windows-1252, ISO-8859-1, etc.) o con caracteres corruptos (mojibake).
2. **Corrección**: Convertir archivos detectados a UTF-8.
3. **Reporte**: Generar un informe de los archivos modificados.

## Uso Previsto
El skill se podrá invocar para escanear un directorio y corregir automáticamente o preguntar antes de cambiar.

## Estructura
- `index.js`: Punto de entrada de la herramienta.
