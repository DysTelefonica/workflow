# Skills

Este repositorio contiene herramientas y "skills" para potenciar el desarrollo con IAs como Trae u OpenCode.

## access-vba-sync

Herramienta para sincronizar bidireccionalmente código VBA (Access) con archivos locales versionables en Git, y generar documentación de base de datos.

**Características:**
- Sincronización bidireccional (Exportar VBA a archivos / Importar archivos a VBA).
- "Watch" mode para desarrollo en tiempo real (edita en VS Code, se actualiza en Access).
- Generación de diagramas ERD / Diccionario de datos en Markdown.
- Manejo de codificación (evita problemas de acentos/mojibake).

### Requisitos
- Windows
- Microsoft Access instalado
- Node.js 18+

### Instalación y Uso

#### Opción A: Uso directo (Standalone)

1. **Instalar dependencias:**
   ```powershell
   cd access-vba-sync
   npm install
   ```

2. **Comandos básicos:**
   ```powershell
   # Iniciar modo Watch (desarrollo en vivo)
   node access-vba-sync/cli.js watch --access "C:\Ruta\TuBaseDeDatos.accdb"

   # Generar documentación ERD
   node access-vba-sync/cli.js generate-erd --backend "C:\Ruta\TuBackend_Datos.accdb"

   # Ayuda completa
   node access-vba-sync/cli.js --help
   ```

#### Opción B: Integración con Agentes (Trae / OpenCode)

Si utilizas un agente de IA, puedes añadir este skill a tu proyecto para que el agente pueda leer y modificar tu código VBA de forma nativa.

1. **Estructura de carpetas:**
   Crea una carpeta `skill` en la raíz de tu proyecto y coloca `access-vba-sync` dentro.
   `MiProyecto/skill/access-vba-sync/SKILL.md`

2. **Uso:**
   El agente detectará automáticamente las capacidades (leer código, exportar, generar ERD) a través del archivo `SKILL.md` y podrá ejecutar las herramientas por ti.

---

## cadete-devops

Skill especializado para el flujo de trabajo con contenedores y OpenShift en entornos corporativos de Telefónica.

**Características:**
- Gestión inteligente de proxy corporativo (VPN/no-VPN).
- Construcción local con Podman y subida a Quay corporativo.
- Despliegue automatizado en OpenShift (Pre-producción y Producción).
- Operaciones de base de datos (Importación SQL y gestión de permisos GRANT).
- Sincronización de volúmenes persistentes (PVC).

### Requisitos
- Windows
- Podman
- oc CLI (OpenShift)

### Instalación y Uso
```powershell
# Ver ayuda completa
python cadete-devops/cli.py --help
```

---

## Cómo instalar en un proyecto nuevo (Sparse Checkout)

Si quieres incorporar `access-vba-sync` a un repositorio existente sin clonar todo el historial de este repo de skills:

```powershell
# Estando en la raíz de tu proyecto:

# 1. Clonar temporalmente sin descargar archivos
git clone --depth 1 --filter=blob:none --sparse https://github.com/ardelperal/skills.git .skills-temp

# 2. Seleccionar solo access-vba-sync
cd .skills-temp
git sparse-checkout set access-vba-sync

# 3. Mover a tu carpeta de skills y limpiar
if (-not (Test-Path "skill")) { New-Item -ItemType Directory -Force -Path "skill" }
Move-Item access-vba-sync ..\skill\access-vba-sync
cd ..
Remove-Item -Recurse -Force .skills-temp

# 4. Instalar dependencias
cd skill/access-vba-sync
npm install
```
