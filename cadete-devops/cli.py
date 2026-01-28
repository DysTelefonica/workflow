#!/usr/bin/env python3
"""
cli.py — Interfaz de línea de comandos para cadete-devops

Uso:
    python cli.py <comando> [opciones]

Ejemplos:
    python cli.py build --version 10.15
    python cli.py push --version 10.15
    python cli.py deploy --cluster pre --ns wcdy-inte-frt --version 10.15
"""

import argparse
import sys
from pathlib import Path

# Añadir el directorio del skill al path
SKILL_DIR = Path(__file__).parent.absolute()
sys.path.insert(0, str(SKILL_DIR))

import oc_manager as oc


def cmd_build(args):
    """Construye una imagen con Podman."""
    version = f"RHL-10.{args.version}" if not args.version.startswith("RHL") else args.version
    image_name = args.image or "cadete"
    dockerfile = args.dockerfile or "Dockerfile"
    context = args.context or "."
    
    success = oc.build_image(dockerfile, image_name, version, context)
    return 0 if success else 1


def cmd_push(args):
    """Sube una imagen a Quay."""
    version = f"RHL-10.{args.version}" if not args.version.startswith("RHL") else args.version
    image_name = args.image or "cadete"
    
    success = oc.push_image(image_name, version)
    return 0 if success else 1


def cmd_deploy(args):
    """Despliega una nueva versión en OpenShift."""
    version = f"RHL-10.{args.version}" if not args.version.startswith("RHL") else args.version
    image_name = args.image or "cadete"
    deployment = args.deployment or image_name
    
    success = oc.update_deployment(deployment, image_name, version, args.cluster, args.ns)
    return 0 if success else 1


def cmd_full_deploy(args):
    """Ejecuta el flujo completo: build + push + deploy."""
    version = f"RHL-10.{args.version}" if not args.version.startswith("RHL") else args.version
    image_name = args.image or "cadete"
    dockerfile = args.dockerfile or "Dockerfile"
    context = args.context or "."
    deployment = args.deployment or image_name
    
    print("=" * 60)
    print("PASO 1: BUILD (Asegúrate de estar FUERA de la VPN)")
    print("=" * 60)
    
    if not oc.build_image(dockerfile, image_name, version, context):
        return 1
    
    print("\n" + "=" * 60)
    print("PASO 2: PUSH (Asegúrate de estar DENTRO de la VPN)")
    print("=" * 60)
    input("\n⏸️  Presiona ENTER cuando estés conectado a la VPN...")
    
    if not oc.push_image(image_name, version):
        return 1
    
    print("\n" + "=" * 60)
    print("PASO 3: DEPLOY")
    print("=" * 60)
    
    if not oc.update_deployment(deployment, image_name, version, args.cluster, args.ns):
        return 1
    
    print("\n" + "=" * 60)
    print("✅ DESPLIEGUE COMPLETO")
    print("=" * 60)
    return 0


def cmd_status(args):
    """Muestra el estado del namespace."""
    oc.get_status(args.cluster, args.ns)
    return 0


def cmd_logs(args):
    """Muestra los logs de un pod."""
    oc.get_pod_logs(args.cluster, args.ns, args.pod, args.tail, args.follow)
    return 0


def cmd_scale(args):
    """Escala un deployment."""
    deployment = args.deployment or "cadete"
    success = oc.scale_deployment(args.cluster, args.ns, deployment, args.replicas)
    return 0 if success else 1


def cmd_rollback(args):
    """Hace rollback de un deployment."""
    deployment = args.deployment or "cadete"
    success = oc.rollback_deployment(args.cluster, args.ns, deployment)
    return 0 if success else 1


def cmd_import_sql(args):
    """Importa un archivo SQL al pod MySQL y configura permisos para 'user'."""
    success = oc.import_sql(args.cluster, args.ns, args.file, args.pod)
    return 0 if success else 1


def cmd_grant_user(args):
    """Solo configura permisos GRANT ALL para el usuario 'user'."""
    success = oc.import_sql(args.cluster, args.ns, None, args.pod)
    return 0 if success else 1


def cmd_sync_pvc(args):
    """Sincroniza archivos locales con un PVC."""
    success = oc.sync_pvc(
        args.cluster, args.ns, args.local,
        args.pvc, args.pod, args.remote or "/data"
    )
    return 0 if success else 1


def cmd_deploy_backup(args):
    """Despliega un backup RAR en todos los entornos."""
    # Ruta por defecto especificada por el usuario
    default_path = r"c:\Proyectos\ubi9\cadete\scripts\backups\ext.rar"
    rar_path = args.file or default_path
    
    success = oc.deploy_backup(rar_path, target_cluster=args.cluster)
    return 0 if success else 1


def cmd_login(args):
    """Login en un cluster de OpenShift."""
    success = oc.oc_login(args.cluster)
    return 0 if success else 1


def main():
    parser = argparse.ArgumentParser(
        description="Cadete-DevOps: Gestión de Quay y OpenShift",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  %(prog)s build --version 10.15
  %(prog)s push --version 10.15
  %(prog)s deploy --cluster pre --ns wcdy-inte-frt --version 10.15
  %(prog)s full-deploy --cluster pre --ns wcdy-inte-frt --version 10.15
  %(prog)s status --cluster pre --ns wcdy-inte-frt
  %(prog)s logs --cluster pre --ns wcdy-inte-frt --pod cadete-xxx
  %(prog)s import-sql --cluster pre --ns wcdy-inte-frt --file backup.sql

Variables de entorno requeridas:
  OC_TOKEN_PRE  Token de OpenShift para preproducción
  OC_TOKEN_PRO  Token de OpenShift para producción
"""
    )
    
    subparsers = parser.add_subparsers(dest="command", help="Comandos disponibles")
    
    # === BUILD ===
    p_build = subparsers.add_parser("build", help="Construir imagen con Podman")
    p_build.add_argument("--version", "-v", required=True, help="Versión (ej: 10.15 o RHL-10.15)")
    p_build.add_argument("--image", "-i", default="cadete", help="Nombre de imagen (default: cadete)")
    p_build.add_argument("--dockerfile", "-f", default="Dockerfile", help="Ruta al Dockerfile")
    p_build.add_argument("--context", "-c", default=".", help="Contexto de build")
    p_build.set_defaults(func=cmd_build)
    
    # === PUSH ===
    p_push = subparsers.add_parser("push", help="Subir imagen a Quay")
    p_push.add_argument("--version", "-v", required=True, help="Versión")
    p_push.add_argument("--image", "-i", default="cadete", help="Nombre de imagen")
    p_push.set_defaults(func=cmd_push)
    
    # === DEPLOY ===
    p_deploy = subparsers.add_parser("deploy", help="Desplegar en OpenShift")
    p_deploy.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_deploy.add_argument("--ns", required=True, help="Namespace")
    p_deploy.add_argument("--version", "-v", required=True, help="Versión")
    p_deploy.add_argument("--image", "-i", default="cadete", help="Nombre de imagen")
    p_deploy.add_argument("--deployment", "-d", help="Nombre del deployment")
    p_deploy.set_defaults(func=cmd_deploy)
    
    # === FULL-DEPLOY ===
    p_full = subparsers.add_parser("full-deploy", help="Build + Push + Deploy completo")
    p_full.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_full.add_argument("--ns", required=True, help="Namespace")
    p_full.add_argument("--version", "-v", required=True, help="Versión")
    p_full.add_argument("--image", "-i", default="cadete", help="Nombre de imagen")
    p_full.add_argument("--dockerfile", "-f", default="Dockerfile", help="Ruta al Dockerfile")
    p_full.add_argument("--context", "-c", default=".", help="Contexto de build")
    p_full.add_argument("--deployment", "-d", help="Nombre del deployment")
    p_full.set_defaults(func=cmd_full_deploy)
    
    # === STATUS ===
    p_status = subparsers.add_parser("status", help="Estado del namespace")
    p_status.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_status.add_argument("--ns", required=True, help="Namespace")
    p_status.set_defaults(func=cmd_status)
    
    # === LOGS ===
    p_logs = subparsers.add_parser("logs", help="Ver logs de un pod")
    p_logs.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_logs.add_argument("--ns", required=True, help="Namespace")
    p_logs.add_argument("--pod", "-p", required=True, help="Nombre del pod")
    p_logs.add_argument("--tail", "-t", type=int, default=100, help="Número de líneas")
    p_logs.add_argument("--follow", "-f", action="store_true", help="Seguir logs en tiempo real")
    p_logs.set_defaults(func=cmd_logs)
    
    # === SCALE ===
    p_scale = subparsers.add_parser("scale", help="Escalar deployment")
    p_scale.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_scale.add_argument("--ns", required=True, help="Namespace")
    p_scale.add_argument("--replicas", "-r", type=int, required=True, help="Número de réplicas")
    p_scale.add_argument("--deployment", "-d", default="cadete", help="Deployment")
    p_scale.set_defaults(func=cmd_scale)
    
    # === ROLLBACK ===
    p_rollback = subparsers.add_parser("rollback", help="Rollback de deployment")
    p_rollback.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_rollback.add_argument("--ns", required=True, help="Namespace")
    p_rollback.add_argument("--deployment", "-d", default="cadete", help="Deployment")
    p_rollback.set_defaults(func=cmd_rollback)
    
    # === IMPORT-SQL ===
    p_sql = subparsers.add_parser("import-sql", help="Importar SQL al pod MySQL (siempre ejecuta GRANT para 'user')")
    p_sql.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_sql.add_argument("--ns", required=True, help="Namespace")
    p_sql.add_argument("--file", "-f", help="Archivo SQL (opcional, si no se pasa solo hace GRANT)")
    p_sql.add_argument("--pod", "-p", help="Pod MySQL (auto-detecta si no se especifica)")
    p_sql.set_defaults(func=cmd_import_sql)
    
    # === GRANT-USER ===
    p_grant = subparsers.add_parser("grant-user", help="Configurar GRANT ALL para usuario 'user' en MySQL")
    p_grant.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_grant.add_argument("--ns", required=True, help="Namespace")
    p_grant.add_argument("--pod", "-p", help="Pod MySQL (auto-detecta si no se especifica)")
    p_grant.set_defaults(func=cmd_grant_user)
    
    # === SYNC-PVC ===
    p_sync = subparsers.add_parser("sync-pvc", help="Sincronizar archivos con PVC")
    p_sync.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_sync.add_argument("--ns", required=True, help="Namespace")
    p_sync.add_argument("--local", "-l", required=True, help="Directorio local")
    p_sync.add_argument("--pvc", help="Nombre del PVC")
    p_sync.add_argument("--pod", "-p", help="Pod con PVC montado")
    p_sync.add_argument("--remote", "-r", help="Ruta remota en el pod")
    p_sync.set_defaults(func=cmd_sync_pvc)
    
    # === DEPLOY-BACKUP ===
    p_backup = subparsers.add_parser("deploy-backup", help="Desplegar backup RAR")
    p_backup.add_argument("--file", "-f", help="Archivo RAR (opcional)")
    p_backup.add_argument("--cluster", "-c", choices=["pre", "pro"], help="Cluster específico (opcional)")
    p_backup.set_defaults(func=cmd_deploy_backup)
    
    # === LOGIN ===
    p_login = subparsers.add_parser("login", help="Login en OpenShift")
    p_login.add_argument("--cluster", required=True, choices=["pre", "pro"], help="Cluster")
    p_login.set_defaults(func=cmd_login)
    
    # Parsear argumentos
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    # Ejecutar comando
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
