#!/usr/bin/env python3
"""
oc_manager.py — Gestor de operaciones para Quay y OpenShift

Este módulo proporciona funciones para:
- Construcción de imágenes con Podman
- Push a Quay corporativo
- Login y operaciones en OpenShift
- Import SQL, sincronización de PVC, etc.
"""

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Optional, List, Dict, Any

# Ruta al directorio del skill
SKILL_DIR = Path(__file__).parent.absolute()
CONFIG_PATH = SKILL_DIR / "config.json"
QUAY_AUTH_PATH = SKILL_DIR / "resources" / "quay_auth.json"
ENV_PATH = SKILL_DIR / ".env"
ENV_EXAMPLE_PATH = SKILL_DIR / ".env.example"


def load_env_file():
    """Carga variables de entorno desde el archivo .env del skill."""
    if not ENV_PATH.exists():
        if ENV_EXAMPLE_PATH.exists():
            print(f"⚠️  No se encontró {ENV_PATH}")
            print(f"💡 Sugerencia: Copia .env.example a .env y configura tus tokens:")
            print(f"   Copy-Item '{ENV_EXAMPLE_PATH}' '{ENV_PATH}'")
        return

    print(f"📄 Cargando configuración desde {ENV_PATH}")
    try:
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                
                if "=" in line:
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()
                    
                    # Solo cargar si tiene valor y no está ya definida
                    if value and key not in os.environ:
                        os.environ[key] = value
    except Exception as e:
        print(f"⚠️  Error leyendo .env: {e}")

# Cargar variables al importar el módulo
load_env_file()


def load_config() -> Dict[str, Any]:
    """Carga la configuración del skill."""
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_cluster_config(config: Dict, cluster: str) -> Dict:
    """Obtiene la configuración de un cluster específico."""
    if cluster not in config["clusters"]:
        raise ValueError(f"Cluster '{cluster}' no válido. Opciones: {list(config['clusters'].keys())}")
    return config["clusters"][cluster]


def get_token(cluster_config: Dict) -> str:
    """Obtiene el token de OpenShift desde variables de entorno."""
    token_env = cluster_config["token_env"]
    token = os.environ.get(token_env)
    if not token:
        raise EnvironmentError(
            f"Token no encontrado. Define la variable de entorno {token_env}\n"
            f"Ejemplo: $env:{token_env} = 'sha256~tu_token'"
        )
    return token


def set_proxy(config: Dict, enable: bool = True):
    """Configura o desactiva el proxy corporativo."""
    if enable:
        os.environ["HTTP_PROXY"] = config["proxy"]["http"]
        os.environ["HTTPS_PROXY"] = config["proxy"]["https"]
        print(f"✓ Proxy configurado: {config['proxy']['http']}")
    else:
        os.environ.pop("HTTP_PROXY", None)
        os.environ.pop("HTTPS_PROXY", None)
        print("✓ Proxy desactivado")


def run_command(cmd: List[str], capture: bool = False, check: bool = True) -> subprocess.CompletedProcess:
    """Ejecuta un comando del sistema."""
    print(f"$ {' '.join(cmd)}")
    result = subprocess.run(
        cmd,
        capture_output=capture,
        text=True,
        check=check
    )
    return result


def build_image(dockerfile_path: str, image_name: str, version: str, context: str = ".") -> bool:
    """
    Construye una imagen con Podman.
    
    IMPORTANTE: Ejecutar SIN VPN para que Podman pueda descargar dependencias.
    """
    config = load_config()
    full_tag = f"{config['quay_registry']}/{config['repository']}/{image_name}:{version}"
    
    print(f"\n🏗️  Construyendo imagen: {full_tag}")
    print("⚠️  Asegúrate de estar FUERA de la VPN\n")
    
    try:
        run_command([
            "podman", "build",
            "-t", full_tag,
            "-f", dockerfile_path,
            context
        ])
        print(f"\n✅ Imagen construida: {full_tag}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Error construyendo imagen: {e}")
        return False


def push_image(image_name: str, version: str) -> bool:
    """
    Sube una imagen a Quay.
    
    IMPORTANTE: Ejecutar CON VPN y con autenticación configurada.
    """
    config = load_config()
    full_tag = f"{config['quay_registry']}/{config['repository']}/{image_name}:{version}"
    
    print(f"\n📤 Subiendo imagen: {full_tag}")
    print("⚠️  Asegúrate de estar DENTRO de la VPN\n")
    
    # Configurar proxy
    set_proxy(config, enable=True)
    
    # Verificar autenticación
    auth_dest = Path.home() / ".config" / "containers" / "auth.json"
    if not auth_dest.exists():
        print(f"⚠️  Copiando autenticación de Quay a {auth_dest}")
        auth_dest.parent.mkdir(parents=True, exist_ok=True)
        import shutil
        shutil.copy(QUAY_AUTH_PATH, auth_dest)
    
    try:
        run_command(["podman", "push", full_tag])
        print(f"\n✅ Imagen subida: {full_tag}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Error subiendo imagen: {e}")
        return False


def oc_login(cluster: str) -> bool:
    """Login en un cluster de OpenShift."""
    config = load_config()
    cluster_config = get_cluster_config(config, cluster)
    
    print(f"\n🔐 Conectando a {cluster_config['name']}...")
    
    # Configurar proxy
    set_proxy(config, enable=True)
    
    try:
        token = get_token(cluster_config)
        run_command([
            "oc", "login",
            "--token", token,
            "--server", cluster_config["server"],
            "--insecure-skip-tls-verify=true"
        ])
        print(f"✅ Conectado a {cluster_config['name']}")
        return True
    except (subprocess.CalledProcessError, EnvironmentError) as e:
        print(f"❌ Error de login: {e}")
        return False


def set_namespace(namespace: str) -> bool:
    """Cambia al namespace especificado."""
    try:
        run_command(["oc", "project", namespace])
        return True
    except subprocess.CalledProcessError:
        print(f"❌ Namespace '{namespace}' no encontrado o sin acceso")
        return False


def update_deployment(deployment: str, image_name: str, version: str, cluster: str, namespace: str) -> bool:
    """Actualiza un deployment con una nueva imagen."""
    config = load_config()
    full_tag = f"{config['quay_registry']}/{config['repository']}/{image_name}:{version}"
    
    print(f"\n🚀 Actualizando deployment '{deployment}' en {namespace}...")
    
    if not oc_login(cluster):
        return False
    
    if not set_namespace(namespace):
        return False
    
    try:
        # Actualizar imagen del deployment
        run_command([
            "oc", "set", "image",
            f"deployment/{deployment}",
            f"{deployment}={full_tag}"
        ])
        
        # Esperar a que el rollout termine
        print("\n⏳ Esperando rollout...")
        run_command([
            "oc", "rollout", "status",
            f"deployment/{deployment}",
            "--timeout=300s"
        ])
        
        print(f"\n✅ Deployment actualizado correctamente")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Error actualizando deployment: {e}")
        return False


def get_pods(cluster: str, namespace: str, selector: Optional[str] = None) -> List[Dict]:
    """Obtiene la lista de pods."""
    if not oc_login(cluster):
        return []
    
    if not set_namespace(namespace):
        return []
    
    cmd = ["oc", "get", "pods", "-o", "json"]
    if selector:
        cmd.extend(["-l", selector])
    
    try:
        result = run_command(cmd, capture=True)
        data = json.loads(result.stdout)
        return data.get("items", [])
    except (subprocess.CalledProcessError, json.JSONDecodeError):
        return []


def get_pod_logs(cluster: str, namespace: str, pod_name: str, tail: int = 100, follow: bool = False) -> None:
    """Muestra los logs de un pod."""
    if not oc_login(cluster):
        return
    
    if not set_namespace(namespace):
        return
    
    cmd = ["oc", "logs", pod_name, f"--tail={tail}"]
    if follow:
        cmd.append("-f")
    
    try:
        run_command(cmd)
    except subprocess.CalledProcessError as e:
        print(f"❌ Error obteniendo logs: {e}")


def get_status(cluster: str, namespace: str) -> None:
    """Muestra el estado de los recursos en el namespace."""
    if not oc_login(cluster):
        return
    
    if not set_namespace(namespace):
        return
    
    print(f"\n📊 Estado de {namespace}\n")
    
    # Pods
    print("=== PODS ===")
    run_command(["oc", "get", "pods", "-o", "wide"], check=False)
    
    # Deployments
    print("\n=== DEPLOYMENTS ===")
    run_command(["oc", "get", "deployments"], check=False)
    
    # PVCs
    print("\n=== PVC ===")
    run_command(["oc", "get", "pvc"], check=False)


def scale_deployment(cluster: str, namespace: str, deployment: str, replicas: int) -> bool:
    """Escala un deployment."""
    if not oc_login(cluster):
        return False
    
    if not set_namespace(namespace):
        return False
    
    print(f"\n⚖️  Escalando {deployment} a {replicas} réplicas...")
    
    try:
        run_command([
            "oc", "scale",
            f"deployment/{deployment}",
            f"--replicas={replicas}"
        ])
        print(f"✅ Deployment escalado a {replicas} réplicas")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error escalando: {e}")
        return False


def rollback_deployment(cluster: str, namespace: str, deployment: str) -> bool:
    """Hace rollback de un deployment a la versión anterior."""
    if not oc_login(cluster):
        return False
    
    if not set_namespace(namespace):
        return False
    
    print(f"\n⏪ Rollback de {deployment}...")
    
    try:
        run_command([
            "oc", "rollout", "undo",
            f"deployment/{deployment}"
        ])
        print("✅ Rollback completado")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error en rollback: {e}")
        return False


def import_sql(cluster: str, namespace: str, sql_file: Optional[str] = None, mysql_pod: Optional[str] = None) -> bool:
    """
    Importa un archivo SQL al pod de MySQL y garantiza permisos para el usuario 'user'.
    
    SIEMPRE ejecuta GRANT ALL PRIVILEGES para el usuario 'user' sobre todas las bases de datos,
    independientemente de si se importa un archivo SQL o no.
    
    Args:
        cluster: Cluster de OpenShift
        namespace: Namespace
        sql_file: Ruta al archivo SQL local (opcional)
        mysql_pod: Nombre del pod MySQL (si no se especifica, se busca automáticamente)
    """
    if not oc_login(cluster):
        return False
    
    if not set_namespace(namespace):
        return False
    
    # Verificar archivo SQL si se proporciona
    if sql_file:
        sql_path = Path(sql_file)
        if not sql_path.exists():
            print(f"❌ Archivo SQL no encontrado: {sql_file}")
            return False
    
    # Buscar pod de MySQL si no se especifica
    if not mysql_pod:
        try:
            result = run_command(
                ["oc", "get", "pods", "-l", "app=mysql", "-o", "jsonpath={.items[0].metadata.name}"],
                capture=True
            )
            mysql_pod = result.stdout.strip()
            if not mysql_pod:
                print("❌ No se encontró pod de MySQL")
                return False
        except subprocess.CalledProcessError:
            print("❌ Error buscando pod de MySQL")
            return False
    
    try:
        # Importar SQL si se proporciona archivo
        if sql_file:
            sql_path = Path(sql_file)
            print(f"\n📥 Importando SQL a pod {mysql_pod}...")
            
            # Copiar archivo al pod
            remote_path = f"/tmp/{sql_path.name}"
            run_command(["oc", "cp", str(sql_path), f"{mysql_pod}:{remote_path}"])
            
            # Ejecutar importación
            run_command([
                "oc", "exec", mysql_pod, "--",
                "bash", "-c",
                f"mysql -u root < {remote_path}"
            ])
            
            # Limpiar archivo temporal
            run_command(["oc", "exec", mysql_pod, "--", "rm", remote_path], check=False)
            
            print("✅ SQL importado correctamente")
        
        # SIEMPRE ejecutar GRANT para el usuario 'user'
        print(f"\n🔑 Configurando permisos para usuario 'user'...")
        
        grant_sql = "GRANT ALL PRIVILEGES ON *.* TO 'user'@'%' WITH GRANT OPTION; FLUSH PRIVILEGES;"
        run_command([
            "oc", "exec", mysql_pod, "--",
            "bash", "-c",
            f"mysql -u root -e \"{grant_sql}\""
        ])
        
        print("✅ Usuario 'user' tiene GRANT ALL en todas las bases de datos")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error: {e}")
        return False


def sync_pvc(cluster: str, namespace: str, local_dir: str, pvc_name: Optional[str] = None, 
             pod_name: Optional[str] = None, remote_path: str = "/data") -> bool:
    """
    Sincroniza archivos locales con un PVC.
    
    Args:
        cluster: Cluster de OpenShift
        namespace: Namespace
        local_dir: Directorio local con los archivos
        pvc_name: Nombre del PVC
        pod_name: Nombre del pod que tiene montado el PVC
        remote_path: Ruta dentro del pod donde está montado el PVC
    """
    if not oc_login(cluster):
        return False
    
    if not set_namespace(namespace):
        return False
    
    local_path = Path(local_dir)
    if not local_path.exists():
        print(f"❌ Directorio local no encontrado: {local_dir}")
        return False
    
    # Buscar pod con el PVC si no se especifica
    if not pod_name:
        config = load_config()
        pvc_name = pvc_name or config.get("pvc", {}).get("anexos", "pvc-anexos")
        
        try:
            # Buscar pod que tenga el PVC montado
            result = run_command(
                ["oc", "get", "pods", "-o", "json"],
                capture=True
            )
            pods = json.loads(result.stdout).get("items", [])
            
            for pod in pods:
                volumes = pod.get("spec", {}).get("volumes", [])
                for vol in volumes:
                    if vol.get("persistentVolumeClaim", {}).get("claimName") == pvc_name:
                        pod_name = pod["metadata"]["name"]
                        # Buscar el mountPath
                        containers = pod.get("spec", {}).get("containers", [])
                        for container in containers:
                            for mount in container.get("volumeMounts", []):
                                if mount.get("name") == vol.get("name"):
                                    remote_path = mount.get("mountPath", remote_path)
                        break
                if pod_name:
                    break
            
            if not pod_name:
                print(f"❌ No se encontró pod con PVC '{pvc_name}' montado")
                return False
                
        except (subprocess.CalledProcessError, json.JSONDecodeError):
            print("❌ Error buscando pod con PVC")
            return False
    
    print(f"\n📂 Sincronizando {local_dir} → {pod_name}:{remote_path}")
    
    try:
        # Usar rsync si está disponible, sino oc cp
        for item in local_path.iterdir():
            src = str(item)
            dest = f"{pod_name}:{remote_path}/{item.name}"
            run_command(["oc", "cp", src, dest])
            print(f"  ✓ {item.name}")
        
        print("\n✅ Sincronización completada")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error sincronizando: {e}")
        return False


if __name__ == "__main__":
    # Test básico
    print("oc_manager.py cargado correctamente")
    config = load_config()
    print(f"Registry: {config['quay_registry']}")
    print(f"Clusters: {list(config['clusters'].keys())}")
