# Cadete-DevOps

Skill para automatizar el flujo de trabajo de desarrollo con contenedores en el entorno corporativo de Telefónica: construcción con Podman, push a Quay y despliegue en OpenShift.

## Instalación Rápida

```powershell
# 1. Configurar autenticación de Quay
Copy-Item .\resources\quay_auth.json $env:USERPROFILE\.config\containers\auth.json

# 2. Configurar tokens de OpenShift
$env:OC_TOKEN_PRE = "sha256~tu_token_preproduccion"
$env:OC_TOKEN_PRO = "sha256~tu_token_produccion"
```

## Uso Básico

```powershell
# Construir imagen (SIN VPN)
python cli.py build --version 10.15

# Subir a Quay (CON VPN)
python cli.py push --version 10.15

# Desplegar en integración
python cli.py deploy --cluster pre --ns wcdy-inte-frt --version 10.15

# Ver estado
python cli.py status --cluster pre --ns wcdy-inte-frt
```

## Comandos Disponibles

| Comando | Descripción |
|---------|-------------|
| `build` | Construir imagen con Podman |
| `push` | Subir imagen a Quay |
| `deploy` | Desplegar en OpenShift |
| `full-deploy` | Flujo completo (build+push+deploy) |
| `status` | Ver estado del namespace |
| `logs` | Ver logs de un pod |
| `scale` | Escalar réplicas |
| `rollback` | Rollback del deployment |
| `import-sql` | Importar SQL a MySQL |
| `sync-pvc` | Sincronizar PVC de anexos |
| `login` | Login en OpenShift |

## Ayuda

```powershell
python cli.py --help
python cli.py <comando> --help
```

## Documentación

Ver [SKILL.md](SKILL.md) para documentación completa.
