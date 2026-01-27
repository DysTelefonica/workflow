# Guía para compartir el Skill Cadete-DevOps

Copia y pega el siguiente texto en un correo para que un compañero pueda instalar el skill en su proyecto.

---

**Asunto: Skill Cadete-DevOps: Automatización de despliegues en Quay y OpenShift**

Hola,

Te paso el "skill" que he preparado para automatizar todo el flujo de trabajo del proyecto **Cadete** con contenedores y OpenShift. Si usas **Trae 2.0** o **Antigravity**, esto permite que la IA actúe como un especialista en nuestra infraestructura corporativa.

### 📋 ¿Qué hace este skill?
- **Construye** las imágenes con Podman (gestiona el proxy corporativo).
- **Sube** las imágenes a nuestro Quay (`wcdy`).
- **Despliega** en OpenShift (Preproducción: integración/certificación y Producción).
- **Operaciones extra**: Logs de pods, importación de SQL a MySQL con permisos automáticos (`GRANT`), y sincronización de anexos en el PVC.

---

### 🚀 Pasos para instalarlo en tu proyecto

**1. Descarga el skill en tu repositorio**
Desde la raíz de tu proyecto, ejecuta estos comandos en PowerShell:

```powershell
# Bajamos solo la carpeta necesaria del repo de skills
git clone --depth 1 --filter=blob:none --sparse https://github.com/ardelperal/skills.git .skills-temp
cd .skills-temp
git sparse-checkout set cadete-devops
cd ..

# Lo movemos a tu carpeta de skills
if (-not (Test-Path "skill")) { New-Item -ItemType Directory -Path "skill" }
Move-Item .skills-temp\cadete-devops .\skill\cadete-devops
Remove-Item -Recurse -Force .skills-temp
```

**2. Configuración de credenciales (Solo la primera vez)**
Para que el motor de contenedores pueda subir imágenes a nuestro Quay:
```powershell
Copy-Item .\skill\cadete-devops\resources\quay_auth.json $env:USERPROFILE\.config\containers\auth.json
```

**3. Configuración de tokens de OpenShift**
Añade tus tokens temporales a la sesión de PowerShell (o a tu perfil de Windows):
```powershell
$env:OC_TOKEN_PRE = "sha256~TU_TOKEN_PRE"
$env:OC_TOKEN_PRO = "sha256~TU_TOKEN_PRO"
```

---

### 💡 Cómo usarlo

**Con la IA (Trae 2.0 / Antigravity):**
Simplemente pídele cosas en lenguaje natural:
- *"Construye y sube la imagen cadete versión 10.15"*
- *"Despliega en el namespace de certificación"*
- *"Configura los permisos del usuario 'user' en el pod de MySQL"*

**Por línea de comandos:**
```powershell
# Nota: Sal de la VPN para hacer el BUILD
python skill\cadete-devops\cli.py build --version 10.15

# Nota: Entra en la VPN para el PUSH y DEPLOY
python skill\cadete-devops\cli.py push --version 10.15
python skill\cadete-devops\cli.py deploy --cluster pre --ns wcdy-inte-frt --version 10.15
```

Cualquier duda me dices.

Un saludo,
