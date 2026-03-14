# Plan de implementación — Framework `workflow`
## Aprovisionamiento automático desde Git para cualquier proyecto

Repositorio base del framework:

https://github.com/DysTelefonica/workflow.git

Objetivo:  
Que cualquier compañero pueda crear o convertir un proyecto en **proyecto compatible con el sistema workflow** ejecutando **un único comando**, sin tener el repo `workflow` previamente clonado.

---

# 1. Objetivo del sistema

El framework `workflow` debe permitir:

- Crear proyectos Access con estructura SDD
- Instalar automáticamente:
  - skills
  - reglas IA
  - plantillas
  - CLI
  - sincronización VBA
- Exportar automáticamente los módulos VBA del Access
- Integrarse con Git
- Preparar entorno para IA (PRD + Specs + ERD)

---

# 2. Modelo de uso final (experiencia del desarrollador)

Crear proyecto nuevo:


mkdir condor
cd condor

npx github:DysTelefonica/workflow init access


o cuando esté publicado en npm:


npx @dys/workflow init access


Resultado:


✔ estructura creada
✔ skills instaladas
✔ reglas copiadas
✔ templates creadas
✔ Access export realizado
✔ proyecto listo para trabajar


---

# 3. Arquitectura del repo workflow

El repositorio debe reorganizarse así:


workflow
│
├─ cli
│ workflow.js
│
├─ installers
│ init-access.js
│ init-project.js
│
├─ templates
│ AGENTS_template.md
│ project_context_template.md
│
├─ skills
│ access-vba-sync
│ prd-writer
│ spec-writer
│ sdd-protocol
│
├─ rules
│ user_rules.md
│ engram-memory-quality.md
│
├─ scripts
│
├─ package.json
│
└─ README.md


---

# 4. CLI principal del framework

Archivo:


cli/workflow.js


Debe ofrecer comandos:


workflow init access
workflow init service
workflow init tool

workflow spec new
workflow release
workflow hotfix

workflow access start
workflow access watch
workflow access erd


---

# 5. Instalador principal

Archivo:


installers/init-access.js


Este script debe realizar:

### 1. Crear estructura de proyecto


docs/
docs/specs/
docs/specs/active/
docs/specs/completed/
docs/PRD/

src/
src/modules/
src/classes/
src/forms/

data/
skills/
rules/


---

### 2. Copiar assets del framework

Copiar desde el repo workflow:


templates/
skills/
rules/


hacia el proyecto.

---

### 3. Crear archivos base

Generar:


AGENTS.md
project_context.md
.gitignore


usando templates.

---

### 4. Detectar Access DB

Buscar en root:


*.accdb
*.mdb
*.accde


Si hay varias:

- elegir determinista
- avisar al usuario.

---

### 5. Instalar skill Access

El framework debe ejecutar:


npm install ./skills/access-vba-sync


o usar link local.

---

### 6. Export inicial del código VBA

Ejecutar automáticamente:


access-vba-sync start


Resultado:


src/modules
src/classes
src/forms


---

### 7. Generar ERD inicial

Ejecutar:


access-vba-sync generate-erd


Resultado:


docs/structure.md


---

# 6. Estructura final de un proyecto Access

Después de ejecutar el instalador:


condor
│
├─ src
│ modules
│ classes
│ forms
│
├─ docs
│ PRD
│ specs
│ active
│ completed
│
├─ rules
│
├─ skills
│
├─ data
│
├─ AGENTS.md
│
├─ project_context.md
│
└─ package.json


---

# 7. Flujo de trabajo del desarrollador

### export inicial


access-vba-sync start


---

### sincronización automática


access-vba-sync watch


---

### generar ERD


access-vba-sync generate-erd


---

### cerrar sesión


access-vba-sync end


---

# 8. Integración con IA

El framework deja el proyecto preparado para:


sdd-protocol


flujo:


User request
↓
Spec generation
↓
Branch creation
↓
Code changes
↓
Access sync
↓
Compile
↓
Release


---

# 9. Distribución del framework

Hay tres opciones.

### opción 1 — ejecutar desde GitHub


npx github:DysTelefonica/workflow init access


Ventajas:

- no requiere instalación
- siempre última versión

---

### opción 2 — instalar global


npm install -g @dys/workflow


uso:


workflow init access


---

### opción 3 — usar CLI local


git clone https://github.com/DysTelefonica/workflow

npm install
npm link


---

# 10. Fases de implementación

## Fase 1
Reorganizar repo `workflow`

- mover scripts a `installers`
- separar CLI
- limpiar templates

---

## Fase 2
Crear instalador


init-access.js


---

## Fase 3
Integrar CLI


workflow init access


---

## Fase 4
Probar en proyectos reales

- CONDOR
- BRASS
- HPS

---

## Fase 5
Publicar CLI


npm publish


o usar GitHub directamente.

---

# 11. Mejoras futuras

Comandos adicionales:


workflow create access-app
workflow create service
workflow create tool


---

Integración adicional:


workflow access compile
workflow access erd
workflow access export
workflow access sync


---

# 12. Resultado esperado

El framework permitirá que cualquier compañero cree un proyecto completo con:


1 comando


y tenga inmediatamente:


Access source control
IA workflow
PRD
Specs
ERD
Git integration


---

# Conclusión

El repositorio `workflow` se convertirá en:

**framework de desarrollo moderno para proyectos Microsoft Access + IA**.

Permitiendo replicar entornos de desarrollo completos en segundos.