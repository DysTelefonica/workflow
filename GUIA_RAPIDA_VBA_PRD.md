# VBA Access PRD Generator - Quick Start Guide

## Overview

Este skill te ayuda a generar PRDs (Product Requirements Documents) exhaustivos para tus aplicaciones VBA Access que quieres migrar a tecnologías web modernas.

## Prerequisitos

Antes de usar este skill, necesitas:

1. **Código fuente exportado**: Tu base de datos Access debe estar exportada usando `msaccess-vcs-addin` u otra herramienta similar
   - Los archivos deben estar en una carpeta `src/`
   - Formatos esperados: `Form_*.cls`, `*.bas` (módulos), `*.cls` (clases)

2. **Diagrama ERD**: Un archivo con la estructura de tu base de datos
   - Puede ser: archivo de texto, imagen, SQL DDL, o cualquier representación del esquema

3. **Disponibilidad para interacción**: El proceso es interactivo y te pedirá:
   - Capturas de pantalla de formularios
   - Aclaraciones sobre funcionalidades
   - Confirmación de elementos no utilizados

## Instalación del Skill

1. Descarga el archivo `vba-access-prd-generator.skill`
2. En Claude, sube el archivo (arrastra y suelta o usa el botón de subir archivo)
3. El skill se instalará automáticamente

## Cómo Usar

### Paso 1: Iniciar el Proceso

Simplemente di:
```
"Quiero generar un PRD para mi aplicación Access usando el skill de VBA PRD Generator"
```

Claude activará el skill y te hará las preguntas iniciales:
- Nombre de la aplicación
- Propósito principal
- Usuarios objetivo
- Ubicación de los archivos fuente
- Ubicación del ERD

### Paso 2: Proporciona los Archivos

Sube a Claude:
- La carpeta `src/` (comprimida en ZIP si es necesario)
- El archivo ERD

### Paso 3: Proceso Interactivo

Claude analizará tu código y luego te guiará formulario por formulario:

**Para cada formulario te pedirá:**
- Captura de pantalla (si el formulario es complejo visualmente)
- Propósito del formulario desde la perspectiva del usuario
- Confirmación de reglas de negocio no obvias en el código
- Confirmación de si el formulario sigue en uso

**Ejemplo de interacción:**
```
Claude: "He analizado el formulario 'Form_FormClientes' y he detectado:
- 15 controles (textboxes, combos, botones)
- Conexión a la tabla TbClientes
- Botones de Guardar, Cancelar, Buscar

¿Puedes compartir una captura de pantalla de este formulario?
¿Cuál es el propósito principal desde el punto de vista del usuario?"

Tú: [Subes screenshot]
"Este formulario permite al usuario de ventas registrar nuevos clientes 
y modificar datos de clientes existentes"

Claude: "Perfecto. He inferido las siguientes user stories:
1. Como usuario de ventas, quiero registrar nuevos clientes...
2. Como usuario de ventas, quiero buscar clientes existentes...
¿Son correctas estas user stories?"
```

### Paso 4: Revisión de Resultados

Al finalizar, Claude generará:

1. **PRD Master** (`PRD_Master_[AppName].md`)
   - Resumen ejecutivo
   - Arquitectura general
   - Modelo de datos completo
   - Estrategia de migración

2. **PRDs por Módulo** (`modules/PRD_Module_[FormName].md`)
   - Documentación detallada de cada formulario
   - User stories específicas
   - Especificaciones de API
   - Consideraciones de migración

3. **Diagramas** (`diagrams/`)
   - Flujo de navegación (Mermaid)
   - Dependencias entre módulos
   - Modelo de datos visual

4. **Documentos Técnicos** (`technical/`)
   - Todas las especificaciones de API
   - Todas las user stories consolidadas
   - Reporte de calidad de código
   - Código muerto identificado

## Características Principales

### ✅ Análisis Exhaustivo
- Escanea todos los formularios, módulos y clases
- Extrae controles, eventos, funciones y procedimientos
- Identifica consultas SQL y accesos a tablas
- Detecta llamadas a APIs externas

### ✅ Detección de Código Muerto
- Identifica formularios que nunca se abren
- Detecta funciones no utilizadas
- Sugiere código candidato para eliminación

### ✅ Mapeo de Navegación
- Genera diagramas visuales de flujo entre formularios
- Identifica puntos de entrada de la aplicación
- Detecta formularios huérfanos

### ✅ Inferencia Inteligente
- Genera user stories basadas en funcionalidad detectada
- Propone endpoints de API REST necesarios
- Sugiere validaciones y reglas de negocio

### ✅ Recomendaciones de Stack
- Sugiere tecnologías backend apropiadas
- Recomienda frameworks frontend
- Propone estrategia de migración por fases

## Ejemplo de Salida

### User Story Generada
```markdown
**Story: Registro de Nueva No Conformidad**

As a quality manager
I want to register a new non-conformity
So that I can track and manage quality issues

Acceptance Criteria:
- [ ] Required fields: Código, Descripción, Fecha Apertura, Responsable
- [ ] Validation: Código debe ser único
- [ ] Auto-generate código based on year and sequence
- [ ] Save creates entry in TbNoConformidades with Estado = "Abierta"

Technical Notes:
- VBA Implementation: Form_Load loads combos, btnGuardar validates and inserts
- Data Operations: INSERT into TbNoConformidades
- Modern Equivalent: POST /api/no-conformidades with JSON payload
```

### API Endpoint Especificado
```markdown
#### POST /api/no-conformidades

**Purpose:** Create a new non-conformity record
**VBA Origin:** btnGuardar_Click in Form_FormNoConformidad

**Request Body:**
```json
{
  "juridica": "string",
  "proyecto": "string",
  "descripcion": "string",
  "fechaApertura": "date",
  "responsable": "string",
  "tipo": "string"
}
```

**Business Rules:**
- Código auto-generated: [YEAR]-[SEQUENCE]
- Estado defaults to "Abierta"
- FechaApertura defaults to current date if not provided
- Validation: Proyecto must exist in TbProyectos
```

## Consejos para Mejores Resultados

1. **Prepara Capturas de Pantalla**: Ten listas las capturas de tus formularios principales
2. **Conoce tu Aplicación**: Familiarízate con los flujos de usuario principales
3. **Identifica Prioridades**: Piensa qué formularios son críticos vs. opcionales
4. **Documenta Excepciones**: Ten claras las reglas de negocio especiales
5. **Revisa el ERD**: Asegúrate que el ERD esté actualizado

## Soporte y Personalización

El skill incluye:
- **Scripts Python** para análisis automatizado
- **Plantillas de PRD** personalizables
- **Patrones de migración** con best practices
- **Referencias de VBA → Web** para cada patrón común

Puedes modificar las plantillas en `references/prd_templates.md` si necesitas un formato específico de PRD para tu organización.

## Próximos Pasos Después del PRD

Una vez tengas tus PRDs:

1. **Prioriza Features**: Usa los PRDs para decidir qué migrar primero
2. **Planifica Sprints**: Convierte user stories en tareas de desarrollo
3. **Diseña APIs**: Usa las especificaciones generadas como base
4. **Mockups UI**: Usa las capturas y descripciones para diseñar nuevas UIs
5. **Plan de Datos**: Usa el análisis de BD para diseñar migraciones

## Troubleshooting

**El análisis falla:**
- Verifica que los archivos `.cls` y `.bas` sean válidos y estén en UTF-8
- Asegúrate que la carpeta `src/` existe y contiene archivos

**No detecta formularios:**
- Los formularios deben empezar con `Form_` en el nombre del archivo
- Verifica que sean archivos `.cls`

**Faltan dependencias:**
- Algunos formularios pueden abrirse dinámicamente por nombre
- El análisis detecta solo `DoCmd.OpenForm` explícitos
- Revisa manualmente formularios que se abren con variables

**El ERD no se procesa bien:**
- Prueba con diferentes formatos (texto plano suele funcionar mejor)
- Puedes copiar/pegar el contenido directamente en el chat

---

¿Tienes dudas? ¡Simplemente pregúntale a Claude mientras usas el skill!
