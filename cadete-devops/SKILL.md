# SKILL.md — Skill Cadete-DevOps para Quay/OpenShift Corporativo

## Objetivo

Este skill automatiza el flujo de trabajo de desarrollo con contenedores para el proyecto **Cadete** en el entorno corporativo de Telefónica:

1. **Construcción** de imágenes Docker con Podman (fuera de VPN)
2. **Push** a Quay corporativo (dentro de VPN)
3. **Despliegue** y gestión en clusters OpenShift de preproducción y producción
4. **Operaciones** adicionales: logs, import SQL, sincronización de PVC, escalado, rollback

---

## Requisitos

| Requisito | Descripción |
|-----------|-------------|
| **Podman** | Instalado y configurado |
| **oc CLI** | Cliente de OpenShift instalado |
| **Python 3.8+** | Para ejecutar los scripts |
| **Acceso VPN** | VPN corporativa para acceso a Quay y OpenShift |

---

## Configuración del Entorno

### Variables de Entorno

```powershell
# Tokens de OpenShift (cambian frecuentemente)
$env:OC_TOKEN_PRE = "sha256~tu_token_preproduccion"
$env:OC_TOKEN_PRO = "sha256~tu_token_produccion"
```

### Autenticación en Quay

El archivo `resources/quay_auth.json` contiene las credenciales de Quay. Cópialo a tu configuración de Podman:

```powershell
# Copiar configuración de autenticación
Copy-Item .\resources\quay_auth.json $env:USERPROFILE\.config\containers\auth.json
```

---

## Flujo de Trabajo Típico

> [!WARNING]
> **VPN**: Debes **SALIR de la VPN** para construir (necesita internet) y **ENTRAR en la VPN** para push/deploy.

### 1. Construir imagen (SIN VPN)

```powershell
python cli.py build --version 10.15
```

### 2. Subir a Quay (CON VPN)

```powershell
python cli.py push --version 10.15
```

### 3. Desplegar en preproducción

```powershell
# Integración
python cli.py deploy --cluster pre --ns wcdy-inte-frt --version 10.15

# Certificación
python cli.py deploy --cluster pre --ns wcdy-cert-frt --version 10.15
```

### 4. Desplegar en producción

```powershell
python cli.py deploy --cluster pro --ns wcdy-prod-frt --version 10.15
```

---

## Comandos Disponibles

### Construcción y Push

| Comando | Descripción |
|---------|-------------|
| `build --version X.XX` | Construir imagen `cadete:RHL-10.XX` |
| `push --version X.XX` | Subir imagen a Quay |
| `full-deploy --version X.XX --cluster pre --ns NS` | Build + Push + Deploy |

### Gestión de OpenShift

| Comando | Descripción |
|---------|-------------|
| `deploy --cluster CLU --ns NS --version X.XX` | Actualizar deployment |
| `status --cluster CLU --ns NS` | Ver estado de pods |
| `logs --cluster CLU --ns NS --pod POD` | Ver logs de un pod |
| `scale --cluster CLU --ns NS --replicas N` | Escalar réplicas |
| `rollback --cluster CLU --ns NS` | Rollback del deployment |

### Operaciones con Datos

| Comando | Descripción |
|---------|-------------|
| `import-sql --cluster CLU --ns NS --file SQL` | Importar SQL al pod MySQL |
| `sync-pvc --cluster CLU --ns NS --local DIR` | Sincronizar PVC de anexos |

---

## Clusters y Namespaces

### Preproducción (`--cluster pre`)

- Server: `https://api.ocgc4pgpre01.mgmt.test.dc.es.telefonica:6443`
- Namespaces:
  - `wcdy-inte-frt` (Integración)
  - `wcdy-cert-frt` (Certificación)

### Producción (`--cluster pro`)

- Server: `https://api.ocgc4pgpro01.mgmt.dc.es.telefonica:6443`
- Namespaces:
  - `wcdy-prod-frt`

---

## Imágenes

| Nombre | Patrón | Base |
|--------|--------|------|
| `cadete` | `RHL-10.XX` | PHP/Apache en UBI9 |
| `mysql` | Sin patrón fijo | MySQL |

---

## Proxy Corporativo

Cuando estés en VPN, el skill configura automáticamente:

```
HTTP_PROXY=http://185.46.212.88:80
HTTPS_PROXY=http://185.46.212.88:80
```

---

## Estructura del Skill

```
cadete-devops/
├── SKILL.md           # Esta documentación
├── README.md          # Guía rápida
├── cli.py             # Interfaz de comandos
├── oc_manager.py      # Lógica de operaciones
├── config.json        # Configuración por defecto
└── resources/
    └── quay_auth.json # Credenciales de Quay
```

---

## Troubleshooting

### Error de conexión al construir

Asegúrate de estar **fuera de la VPN** para que Podman pueda descargar dependencias.

### Error de autenticación en Quay

Verifica que `quay_auth.json` esté copiado a `~/.config/containers/auth.json`.

### Token de OpenShift expirado

Obtén un nuevo token desde la consola web de OpenShift y actualiza la variable de entorno:
```powershell
$env:OC_TOKEN_PRE = "sha256~nuevo_token"
```
