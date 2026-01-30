# VBA Cleaner Skill

Skill para limpiar y arreglar problemas de codificación en archivos VBA exportados de Microsoft Access.

## 📦 Contenido

Este paquete incluye:

- **vba-cleaner.skill** - Archivo comprimido listo para importar en Claude
- **vba-cleaner/** - Directorio sin comprimir del skill (opcional, para desarrollo)

## 🚀 Instalación

### Opción 1: Usar el archivo .skill (Recomendado)

1. Abre Claude.ai
2. Ve a Settings → Skills
3. Haz clic en "Import Skill"
4. Selecciona el archivo `vba-cleaner.skill`

### Opción 2: Copiar el directorio manualmente

Si estás desarrollando o modificando el skill:

1. Copia la carpeta `vba-cleaner/` a tu repositorio de skills
2. Colócala en la ubicación apropiada según tu configuración

## 🎯 ¿Qué hace este skill?

Resuelve automáticamente los problemas comunes al exportar código VBA desde Microsoft Access:

- ✅ Detecta y corrige la codificación del archivo
- ✅ Elimina caracteres BOM (Byte Order Mark) que aparecen como "BOOM"
- ✅ Convierte todo a UTF-8 limpio sin BOM
- ✅ Normaliza los saltos de línea
- ✅ Preserva la estructura y funcionalidad del código VBA

## 💡 Casos de uso

El skill se activará automáticamente cuando:

- Subas archivos VBA (.bas, .cls, .frm)
- Pidas "limpia este archivo VBA"
- Solicites "arregla la codificación de estos módulos"
- Menciones "elimina el BOM" o problemas de encoding en VBA

## 📋 Ejemplos de uso

Una vez instalado el skill, simplemente:

1. **Archivo individual:**
   - Sube tu archivo .bas
   - Di: "Limpia este archivo VBA"

2. **Múltiples archivos:**
   - Sube varios archivos .bas, .cls o .frm
   - Di: "Arregla la codificación de todos estos módulos VBA"

3. **Verificación:**
   - Di: "Este archivo VBA tiene caracteres extraños, ¿puedes arreglarlo?"

## 🔧 Dependencias

El skill requiere la biblioteca Python `chardet` que se instalará automáticamente cuando Claude ejecute el script por primera vez.

## 📄 Tipos de archivo soportados

- `.bas` - Módulos estándar
- `.cls` - Módulos de clase
- `.frm` - Módulos de formulario
- Cualquier archivo de texto con código VBA

## 🛠️ Desarrollo

Si quieres modificar el skill:

1. Edita los archivos en `vba-cleaner/`
2. El script principal está en `vba-cleaner/scripts/clean_vba.py`
3. La documentación está en `vba-cleaner/SKILL.md`
4. Vuelve a empaquetar con el script de packaging si es necesario

## ❓ Soporte

Si tienes problemas:

1. Verifica que el archivo subido sea realmente un archivo VBA de texto
2. Asegúrate de que Claude tiene acceso a ejecutar scripts Python
3. Revisa que la dependencia `chardet` esté instalada

---

**Versión:** 1.0  
**Creado para:** Claude AI  
**Compatibilidad:** Todas las versiones de Claude con capacidad de ejecución de código
