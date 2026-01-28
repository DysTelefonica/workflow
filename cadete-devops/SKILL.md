# SKILL.md — Skill Cadete-DevOps para Quay/OpenShift Corporativo

## 1. Visión General y Objetivo

Este skill encapsula la lógica de operaciones DevOps para el proyecto **Cadete** en la infraestructura de Telefónica. Su propósito es permitir a agentes de IA y desarrolladores humanos realizar tareas complejas de despliegue y mantenimiento de forma estandarizada y segura.

**Capacidades Principales:**
*   **Ciclo de Vida de Imágenes:** Construcción (Podman) y publicación (Quay) de imágenes Docker.
*   **Gestión de Despliegues:** Actualización, escalado, rollback y monitoreo en clusters OpenShift (Preproducción y Producción).
*   **Gestión de Datos:** Sincronización masiva de adjuntos (RAR/PVC), importación de bases de datos SQL y gestión de permisos.
*   **Abstracción de Infraestructura:** Manejo automático de autenticación, proxies corporativos y contextos de seguridad.

---

## 2. Requisitos del Entorno

Para que este skill funcione correctamente, el entorno de ejecución debe cumplir:

| Componente | Requisito | Notas para el Agente |
| :--- | :--- | :--- |
| **Sistema Operativo** | Windows (recomendado) o Linux | El skill maneja rutas de forma agnóstica (`pathlib`), pero está optimizado para Windows (ej. rutas de 7-Zip). |
| **Python** | 3.8+ | Requiere librerías estándar (`subprocess`, `json`, `argparse`). |
| **Container Engine** | Podman | Debe estar en el PATH. Necesario para `build` y `push`. |
| **OpenShift CLI** | `oc` v4.x | Debe estar en el PATH. |
| **Compresión** | 7-Zip | Requerido para `deploy-backup`. Se busca en `C:\Program Files\7-Zip\7z.exe` o en PATH. |
| **Red** | Acceso VPN / Internet | **Build:** Requiere Internet (SIN VPN). **Push/Deploy:** Requiere Intranet (CON VPN). |

---

## 3. Configuración

### 3.1 Variables de Entorno (`.env`)
El skill busca un archivo `.env` en su directorio raíz. Si no existe, usa variables de entorno del sistema.

**Variables Críticas:**
```ini
# Tokens de autenticación (Obtener de consola web OpenShift -> Copy Login Command)
OC_TOKEN_PRE=sha256~...      # Token para clúster de Preproducción
OC_TOKEN_PRO=sha256~...      # Token para clúster de Producción
```

### 3.2 Archivo de Configuración (`config.json`)
Define la topología de los clústeres y los recursos.

```json
{
  "quay_registry": "quay.apps.ocgc4tools.mgmt.dc.es.telefonica",
  "repository": "wcdy",
  "proxy": { ... },
  "clusters": {
    "pre": {
      "name": "Preproducción",
      "server": "https://api.ocgc4pgpre01...",
      "namespaces": ["wcdy-inte-frt", "wcdy-cert-frt"],
      "token_env": "OC_TOKEN_PRE"
    },
    "pro": { ... }
  }
}
```

---

## 4. Guía de Comandos (CLI)

El punto de entrada es `cli.py`. Todos los comandos retornan **Exit Code 0** en éxito y **1** en error.

### 4.1 Ciclo de Imagen (Build & Push)

> [!IMPORTANT]
> **Gestión de Red:** El Agente debe verificar la conectividad antes de ejecutar estos comandos.
> *   `build`: Fallará si estás en VPN (no puede bajar paquetes de repos públicos).
> *   `push`: Fallará si NO estás en VPN (no puede conectar a Quay corporativo).

| Comando | Argumentos Clave | Descripción |
| :--- | :--- | :--- |
| `build` | `--version <V>` | Construye la imagen `cadete:RHL-10.<V>`. |
| `push` | `--version <V>` | Sube la imagen a Quay. Configura proxy automáticamente. |
| `full-deploy` | `--cluster <C>`, `--ns <N>`, `--version <V>` | Ejecuta Build -> Pausa (para conectar VPN) -> Push -> Deploy. |

### 4.2 Operaciones en OpenShift

| Comando | Descripción | Ejemplo |
| :--- | :--- | :--- |
| `deploy` | Actualiza la imagen del deployment. | `python cli.py deploy --cluster pre --ns wcdy-inte-frt --version 10.25` |
| `scale` | Escala el número de réplicas. | `python cli.py scale --cluster pro --ns wcdy-prod-frt --replicas 2` |
| `rollback` | Revierte a la versión anterior. | `python cli.py rollback --cluster pre --ns wcdy-inte-frt` |
| `status` | Muestra pods, deployments y PVCs. | `python cli.py status --cluster pre --ns wcdy-inte-frt` |
| `logs` | Muestra logs (soporta `-f` follow). | `python cli.py logs --cluster pre --ns wcdy-inte-frt --pod <pod-name>` |

### 4.3 Gestión de Datos y Backups

Estas operaciones son críticas para la consistencia de datos entre entornos.

#### `deploy-backup` (Sincronización Masiva de Adjuntos)
Despliega un archivo RAR de adjuntos en **todos** los entornos configurados (o uno específico).

*   **Uso:** `python cli.py deploy-backup --file "c:\ruta\ext.rar" [--cluster pre]`
*   **Lógica Interna:**
    1.  Descomprime el RAR en un directorio temporal local (usa 7-Zip).
    2.  Detecta si existe una subcarpeta `ext/` y la usa como raíz.
    3.  Itera sobre todos los clusters y namespaces definidos en `config.json`.
    4.  **Auto-descubrimiento:** Encuentra el PVC de adjuntos (excluyendo PVCs de SQL).
    5.  **Sincronización:** Empaqueta en TAR y transfiere al pod, descomprimiendo con `--overwrite`. No borra archivos existentes en destino (modo aditivo/actualización).

#### `import-sql` (Restauración de Base de Datos)
Importa un dump SQL en el pod de MySQL.

*   **Uso:** `python cli.py import-sql --cluster pre --ns wcdy-inte-frt --file backup.sql`
*   **Características:**
    *   Limpia `DEFINER` y `LOCK TABLES` del SQL automáticamente para evitar errores de permisos.
    *   Ejecuta `GRANT ALL` para el usuario de aplicación tras la importación.
    *   Soporta detección automática del pod MySQL.

#### `sync-pvc` (Sincronización Genérica)
Sincroniza un directorio local cualquiera con un PVC remoto.

*   **Uso:** `python cli.py sync-pvc --cluster pre --ns wcdy-inte-frt --local ./dist --pvc web-pvc`

---

## 5. Instrucciones para Agentes de IA

Si estás utilizando este skill para cumplir una solicitud de usuario, ten en cuenta:

1.  **Verificación de Éxito:**
    *   Confía en el **Exit Code**. Si es `0`, la operación fue exitosa.
    *   Lee la salida estándar (stdout). El skill imprime emojis (✅, ❌, ⚠️) para indicar estado visualmente.

2.  **Manejo de Errores Comunes:**
    *   **"No such host" / "Dial tcp":** Indica problema de VPN.
        *   Si es `build`: Sugiere desconectar VPN.
        *   Si es `push/deploy`: Sugiere conectar VPN.
    *   **"Token expired" / "Unauthorized":** El token en `.env` ha caducado. Solicita al usuario uno nuevo.
    *   **"Multi-Attach error":** Un volumen RWO está bloqueado. Sugiere escalar a 0 y luego a 1, o cambiar estrategia a `Recreate`.

3.  **Seguridad:**
    *   Nunca imprimas el contenido de `.env` o los tokens en el chat.
    *   Al importar SQL, el skill maneja las credenciales internamente; no necesitas pedirlas.

4.  **Persistencia:**
    *   Si necesitas hacer cambios permanentes en la infraestructura (ej. añadir un nuevo namespace), edita `config.json`.

---

## 6. Estructura de Archivos

```
cadete-devops/
├── cli.py             # Punto de entrada (ArgParse)
├── oc_manager.py      # Lógica de negocio y wrappers de OC/Podman
├── config.json        # Configuración de clusters y endpoints
├── SKILL.md           # Esta documentación
├── README.md          # Resumen rápido
├── manifests/         # (Opcional) Archivos JSON/YAML de referencia
└── resources/
    └── quay_auth.json # Plantilla de credenciales Quay
```
