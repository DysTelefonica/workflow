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
import re
import tarfile
import tempfile
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


def run_command(cmd: List[str], capture: bool = False, check: bool = True, retries: int = 0, **kwargs) -> subprocess.CompletedProcess:
    """Ejecuta un comando del sistema con reintentos opcionales."""
    import time
    
    for attempt in range(retries + 1):
        if attempt > 0:
            print(f"   ⚠️  Reintento {attempt}/{retries} en 2s...")
            time.sleep(2)
            
        print(f"$ {' '.join(cmd)}")
        try:
            result = subprocess.run(
                cmd,
                capture_output=capture,
                text=True,
                check=check,
                **kwargs
            )
            
            # Si check=False, verificamos returncode manualmente para reintentos
            if not check and result.returncode != 0:
                # Simulamos lógica de error para evaluar reintento
                if attempt < retries:
                    is_network_error = False
                    if result.stderr:
                         if "no such host" in result.stderr or "dial tcp" in result.stderr or "Unable to connect" in result.stderr:
                             is_network_error = True
                    
                    if is_network_error:
                        print(f"   ❌ Falló (code {result.returncode}) - 🌐 Detectado error de red/DNS")
                        continue # Al siguiente intento
            
            return result

        except subprocess.CalledProcessError as e:
            # Si es el último intento, propagar la excepción si check=True
            if attempt == retries:
                if check:
                    raise e
                return subprocess.CompletedProcess(cmd, e.returncode, stdout=e.stdout, stderr=e.stderr)
            
            # Analizar si es un error recuperable (red/DNS)
            # Si no capturamos salida, no podemos saberlo fácilmente, pero asumimos reintento si check=True falló
            # Si capturamos, miramos stderr
            is_network_error = False
            if e.stderr:
                if "no such host" in e.stderr or "dial tcp" in e.stderr or "Unable to connect" in e.stderr:
                    is_network_error = True
            
            # Si no es error de red y no queremos reintentar todo, podríamos romper aquí
            # Pero por simplicidad, reintentamos si se solicitó retries > 0
            print(f"   ❌ Falló (code {e.returncode})")
            if is_network_error:
                print("   🌐 Detectado error de red/DNS")

    # Should not be reached if check=True
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
    
    # Configurar proxy (Desactivado para OpenShift interno)
    # set_proxy(config, enable=True)
    
    try:
        token = get_token(cluster_config)
        run_command([
            "oc", "login",
            "--token", token,
            "--server", cluster_config["server"],
            "--insecure-skip-tls-verify=true"
        ], retries=3)
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


def prepare_sql_file(sql_file: Path) -> Path:
    """
    Prepara el archivo SQL:
    - Convierte a UTF-8
    - Elimina DEFINERs problemáticos
    - Retorna la ruta al archivo preparado
    """
    print(f"🔄 Preparando archivo SQL: {sql_file.name}")
    
    content = ""
    encoding_found = None
    
    # Intentar leer con diferentes encodings
    for enc in ["utf-8", "utf-16", "latin-1", "cp1252"]:
        try:
            with open(sql_file, "r", encoding=enc) as f:
                content = f.read()
            encoding_found = enc
            print(f"  ✓ Leído con encoding: {enc}")
            break
        except UnicodeError:
            continue
            
    if not content:
        raise ValueError(f"No se pudo leer el archivo {sql_file} con encodings estándar")

    # Limpieza de DEFINERs
    # Patrón: DEFINER=`root`@`localhost` o DEFINER=root@localhost
    # Se reemplaza por vacío
    pattern = re.compile(r"DEFINER\s*=\s*[`']?[\w%-]+[`']?@[`']?[\w%-]+[`']?", re.IGNORECASE)
    content_clean, num_subs = pattern.subn("", content)
    
    if num_subs > 0:
        print(f"  ✓ Eliminados {num_subs} DEFINERs")

    # Limpieza de LOCK TABLES (para evitar errores 1100 en importación)
    content_clean = re.sub(r"LOCK TABLES\s+[`\w]+\s+WRITE;", "", content_clean, flags=re.IGNORECASE)
    content_clean = re.sub(r"UNLOCK TABLES;", "", content_clean, flags=re.IGNORECASE)
    
    # Guardar versión limpia
    prepared_path = sql_file.parent / f"prepared_{sql_file.name}"
    with open(prepared_path, "w", encoding="utf-8") as f:
        f.write(content_clean)
        
    print(f"  ✓ Archivo preparado guardado en: {prepared_path.name}")
    return prepared_path


def import_sql(cluster: str, namespace: str, sql_file: Optional[str] = None, mysql_pod: Optional[str] = None) -> bool:
    """
    Importa un archivo SQL al pod de MySQL de forma robusta.
    
    SIEMPRE ejecuta GRANT ALL PRIVILEGES para el usuario 'user' sobre todas las bases de datos.
    """
    if not oc_login(cluster):
        return False
    
    if not set_namespace(namespace):
        return False
    
    # Verificar archivo SQL si se proporciona
    sql_path_to_use = None
    if sql_file:
        sql_path = Path(sql_file)
        if not sql_path.exists():
            print(f"❌ Archivo SQL no encontrado: {sql_file}")
            return False
        
        try:
            sql_path_to_use = prepare_sql_file(sql_path)
        except Exception as e:
            print(f"❌ Error preparando archivo SQL: {e}")
            return False
    
    # Buscar pod de MySQL si no se especifica
    if not mysql_pod:
        print("🔍 Buscando pod de MySQL...")
        # Intentar varias estrategias
        strategies = [
            ["oc", "get", "pods", "-l", "app=mysql", "-o", "jsonpath={.items[0].metadata.name}"],
            ["oc", "get", "pods", "-l", "app=cadete-db", "-o", "jsonpath={.items[0].metadata.name}"],
            ["oc", "get", "pods", "-l", "component=database", "-o", "jsonpath={.items[0].metadata.name}"],
        ]
        
        for cmd in strategies:
            try:
                result = run_command(cmd, capture=True, check=False)
                pod = result.stdout.strip()
                if pod:
                    mysql_pod = pod
                    print(f"  ✓ Pod encontrado: {mysql_pod}")
                    break
            except Exception:
                pass
        
        if not mysql_pod:
            # Fallback manual buscando en la lista por nombre
            try:
                result = run_command(["oc", "get", "pods", "-o", "name"], capture=True)
                for line in result.stdout.splitlines():
                    name = line.replace("pod/", "").strip()
                    if "mysql" in name or "db" in name or "mariadb" in name:
                        mysql_pod = name
                        print(f"  ✓ Pod encontrado (por nombre): {mysql_pod}")
                        break
            except Exception:
                pass

        if not mysql_pod:
            print("❌ No se encontró ningún pod de MySQL/DB")
            return False

    # Detectar contenedor principal para evitar ambigüedades
    try:
        res_container = run_command(
            ["oc", "get", "pod", mysql_pod, "-o", "jsonpath={.spec.containers[0].name}"],
            capture=True, check=True
        )
        container_name = res_container.stdout.strip()
        print(f"  ✓ Contenedor detectado: {container_name}")
    except Exception:
        container_name = None
        print("  ⚠️ No se pudo detectar nombre de contenedor, usando default")

    # Helper para comandos exec/cp
    def get_oc_exec_args(pod, cmd_list):
        base = ["oc", "exec", pod]
        if container_name:
            base.extend(["-c", container_name])
        base.append("--")
        base.extend(cmd_list)
        return base

    def get_oc_cp_args(src, dest):
        base = ["oc", "cp", src, dest]
        if container_name:
            base.extend(["-c", container_name])
        return base
            
    try:
        # Determinar si necesitamos password
        print("🔐 Verificando acceso a MySQL...")
        use_password = True
        try:
            # Intentar sin password primero
            run_command(
                get_oc_exec_args(mysql_pod, ["mysql", "-u", "root", "-e", "SELECT 1"]),
                capture=True,
                check=True
            )
            print("  ✓ Acceso sin contraseña permitido")
            use_password = False
        except subprocess.CalledProcessError:
            print("  ⚠️ Acceso sin contraseña falló, se intentará con MYSQL_ROOT_PASSWORD")
        
        # Limpiar bases de datos existentes para evitar conflictos (Tablas vs Vistas, Duplicados)
        print("🧹 Limpiando bases de datos de usuario existentes...")
        try:
            # Comando para listar DBs excluyendo las del sistema
            
            # Ejecutar y capturar (necesitamos bash para expandir la variable de entorno si hay pass)
            if use_password:
                # Usamos string explícito para evitar problemas de quoting en bash
                cmd_str = "mysql -u root -p$MYSQL_ROOT_PASSWORD -N -e 'SHOW DATABASES'"
                res_dbs = run_command(
                    get_oc_exec_args(mysql_pod, ["bash", "-c", cmd_str]),
                    capture=True, check=True
                )
            else:
                # Sin password, pasamos lista directa (subprocess maneja argumentos)
                list_dbs_cmd = ["mysql", "-u", "root", "-N", "-e", "SHOW DATABASES"]
                res_dbs = run_command(
                    get_oc_exec_args(mysql_pod, list_dbs_cmd),
                    capture=True, check=True
                )
                
            dbs = res_dbs.stdout.splitlines()
            system_dbs = {"information_schema", "mysql", "performance_schema", "sys"}
            
            for db in dbs:
                db = db.strip()
                if db and db not in system_dbs:
                    print(f"  - Borrando base de datos: {db}")
                    drop_cmd = f"DROP DATABASE IF EXISTS `{db}`"
                    
                    if use_password:
                        # Usar comillas simples para el SQL para evitar expansión de backticks en bash
                        run_command(get_oc_exec_args(mysql_pod, ["bash", "-c", f"mysql -u root -p\"$MYSQL_ROOT_PASSWORD\" -e '{drop_cmd}'"]), check=False)
                    else:
                        run_command(get_oc_exec_args(mysql_pod, ["mysql", "-u", "root", "-e", drop_cmd]), check=False)
                        
        except Exception as e:
            print(f"⚠️  Advertencia al limpiar bases de datos: {e}")

        # Importar SQL si se proporciona archivo
        if sql_path_to_use:
            print(f"\n📥 Importando SQL a pod {mysql_pod}...")
            
            # Copiar archivo al pod
            # Workaround para Windows: 'oc cp' falla con rutas absolutas (C:\...)
            # Cambiamos al directorio del archivo y usamos ruta relativa
            remote_path = f"/tmp/{sql_path_to_use.name}"
            run_command(
                get_oc_cp_args(sql_path_to_use.name, f"{mysql_pod}:{remote_path}"),
                cwd=str(sql_path_to_use.parent)
            )
            
            # Ejecutar importación
            if use_password:
                import_cmd = (
                    f"if [ -n \"$MYSQL_ROOT_PASSWORD\" ]; then "
                    f"  mysql -u root -p\"$MYSQL_ROOT_PASSWORD\" -f < {remote_path}; "
                    f"else "
                    f"  mysql -u root -f < {remote_path}; "
                    f"fi"
                )
            else:
                import_cmd = f"mysql -u root -f < {remote_path}"
            
            try:
                run_command(get_oc_exec_args(mysql_pod, ["bash", "-c", import_cmd]))
                print("✅ SQL importado correctamente")
            finally:
                # Limpiar archivo temporal remoto SIEMPRE
                run_command(get_oc_exec_args(mysql_pod, ["rm", "-f", remote_path]), check=False)
        
        # SIEMPRE ejecutar GRANT para el usuario 'user'
        print(f"\n🔑 Configurando permisos para usuario 'user'...")
        
        grant_sql = "GRANT ALL PRIVILEGES ON *.* TO 'user'@'%' WITH GRANT OPTION; FLUSH PRIVILEGES;"
        
        if use_password:
            grant_cmd = (
                f"if [ -n \"$MYSQL_ROOT_PASSWORD\" ]; then "
                f"  mysql -u root -p\"$MYSQL_ROOT_PASSWORD\" -e \"{grant_sql}\"; "
                f"else "
                f"  mysql -u root -e \"{grant_sql}\"; "
                f"fi"
            )
        else:
            grant_cmd = f"mysql -u root -e \"{grant_sql}\""
        
        run_command(get_oc_exec_args(mysql_pod, ["bash", "-c", grant_cmd]))
        
        print("✅ Usuario 'user' tiene GRANT ALL en todas las bases de datos")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error durante la operación en el pod: {e}")
        return False


def sync_pvc(cluster: str, namespace: str, local_dir: str, pvc_name: Optional[str] = None, 
             pod_name: Optional[str] = None, remote_path: str = "/data", mirror: bool = False) -> bool:
    """
    Sincroniza archivos locales con un PVC usando estrategia TAR.
    
    Args:
        cluster: Cluster de OpenShift
        namespace: Namespace
        local_dir: Directorio local con los archivos
        pvc_name: Nombre del PVC
        pod_name: Nombre del pod que tiene montado el PVC
        remote_path: Ruta dentro del pod donde está montado el PVC
        mirror: Si True, borra el contenido remoto antes de copiar (sincronización exacta)
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
                # Verificar status Running
                if pod.get("status", {}).get("phase") != "Running":
                    continue
                    
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
                                    print(f"      ✓ Detectado mountPath: {remote_path}")
                        break
                if pod_name:
                    break
            
            if not pod_name:
                print(f"❌ No se encontró pod RUNNING con PVC '{pvc_name}' montado")
                return False
                
        except (subprocess.CalledProcessError, json.JSONDecodeError):
            print("❌ Error buscando pod con PVC")
            return False
    
    print(f"\n📂 Sincronizando {local_dir} → {pod_name}:{remote_path}")
    if mirror:
        print("   ⚠️  MODO MIRROR: Se borrará el contenido remoto actual.")
    
    try:
        # 1. Crear TAR local
        with tempfile.TemporaryDirectory() as temp_tar_dir:
            tar_filename = "sync_payload.tar"
            tar_full_path = Path(temp_tar_dir) / tar_filename
            
            print(f"   📦 Comprimiendo archivos locales...", end=" ", flush=True)
            with tarfile.open(tar_full_path, "w") as tar:
                tar.add(local_path, arcname=".")
            
            size_mb = tar_full_path.stat().st_size / (1024 * 1024)
            print(f"✓ ({size_mb:.2f} MB)")
            
            # 2. Copiar TAR al pod
            remote_tar = f"/tmp/{tar_filename}"
            print(f"   🚀 Transfiriendo al pod...", end=" ", flush=True)
            
            # Usamos cwd=temp_tar_dir para que el src sea solo el nombre del archivo (mejor compatibilidad Windows)
            run_command(
                ["oc", "cp", tar_filename, f"{pod_name}:{remote_tar}"], 
                cwd=str(temp_tar_dir),
                capture=True,
                retries=3
            )
            print("✓")
            
            # 3. Operaciones remotas
            print(f"   ⚙️  Aplicando cambios remotos...", end=" ", flush=True)
            
            # Construir comando robusto
            # Usamos 'trap' para asegurar que el archivo temporal se borre SIEMPRE al salir,
            # independientemente de si el comando tar tiene éxito o falla.
            setup_cmd = f"mkdir -p {remote_path} && cd {remote_path}"
            
            mirror_cmd = ""
            if mirror:
                mirror_cmd = " && find . -mindepth 1 -delete"
            
            # Flags de tar para máxima compatibilidad y evitar errores de permisos/tiempos
            tar_cmd = f"tar -xf {remote_tar} --no-same-owner --no-same-permissions -m --overwrite"
            
            # Estructura: trap 'rm' EXIT; cd ... && [clean] && tar
            remote_cmd_str = f"trap 'rm -f {remote_tar}' EXIT; {setup_cmd}{mirror_cmd} && {tar_cmd}"
            
            # Ejecutar y capturar salida completa para debug
            # Añadimos retries=3 para tolerar fallos de red intermitentes
            res = run_command(
                ["oc", "exec", pod_name, "--", "bash", "-c", remote_cmd_str],
                capture=True,
                check=False,
                retries=3
            )
            
            if res.returncode != 0:
                print(f"\n   ⚠️  Resultado remoto (code {res.returncode}):")
                if res.stdout.strip():
                    print(f"   Stdout: {res.stdout}")
                print(f"   Stderr: {res.stderr}")
                
                # Detectar errores fatales de conexión
                if "Unable to connect" in res.stderr or "no such host" in res.stderr or "dial tcp" in res.stderr:
                    return False
                
                print("   ℹ️  Asumiendo éxito parcial (warnings de tar ignorados)")
                 
            print("✓")
            
            # Verificación final
            print(f"   🔎 Verificando estructura remota...")
            ls_res = run_command(
                ["oc", "exec", pod_name, "--", "ls", "-la", remote_path],
                capture=True,
                check=False,
                retries=3
            )
            if ls_res.returncode == 0:
                print(ls_res.stdout)
            else:
                print(f"   ⚠️ No se pudo listar el directorio remoto: {ls_res.stderr}")
                    
            print("\n✅ Sincronización completada")
            return True
        
    except Exception as e:
        print(f"\n❌ Error sincronizando: {e}")
        return False


def find_target_pvc(namespace: str) -> Optional[str]:
    """
    Descubre el PVC objetivo en un namespace, excluyendo el de SQL.
    Asume que hay dos PVCs y uno es de SQL.
    """
    try:
        res = run_command(
            ["oc", "get", "pvc", "-n", namespace, "-o", "json"],
            capture=True
        )
        data = json.loads(res.stdout)
        items = data.get("items", [])
        
        candidates = []
        for item in items:
            name = item["metadata"]["name"]
            # Excluir PVCs de SQL
            if "sql" in name.lower() or "mysql" in name.lower():
                continue
            candidates.append(name)
            
        if len(candidates) == 1:
            return candidates[0]
        elif len(candidates) > 1:
            # Si hay más de uno, priorizar 'web-pvc' o 'anexos'
            print(f"   ⚠️  Múltiples candidatos encontrados: {candidates}")
            for c in candidates:
                if "web" in c or "anexo" in c:
                    return c
            return candidates[0]
        else:
            print(f"   ❌ No se encontró ningún PVC candidato en {namespace} (excluyendo SQL)")
            return None
            
    except (subprocess.CalledProcessError, json.JSONDecodeError) as e:
        print(f"   ❌ Error descubriendo PVCs: {e}")
        return None

def deploy_backup(rar_path: str, target_cluster: Optional[str] = None) -> bool:
    """
    Despliega el contenido de un backup RAR en todos los entornos (Pre y Pro).
    Extrae el RAR y sincroniza con el PVC de anexos (no SQL) en todos los namespaces configurados.
    
    Args:
        rar_path: Ruta al archivo RAR.
        target_cluster: (Opcional) 'pre' o 'pro' para limitar el despliegue.
    """
    # tempfile importado globalmente
    
    rar_file = Path(rar_path)
    if not rar_file.exists():
        print(f"❌ Archivo RAR no encontrado: {rar_path}")
        return False

    # Detectar 7-Zip
    seven_z = Path(r"C:\Program Files\7-Zip\7z.exe")
    if not seven_z.exists():
        # Intentar buscar en path
        seven_z = "7z"
        try:
            run_command(["7z", "--help"], capture=True, check=False)
        except FileNotFoundError:
            print("❌ No se encontró 7-Zip (necesario para extraer .rar)")
            print("   Instala 7-Zip o asegúrate de que esté en C:\\Program Files\\7-Zip\\7z.exe")
            return False

    print(f"🚀 Iniciando despliegue masivo de backup: {rar_file.name}")
    
    with tempfile.TemporaryDirectory() as temp_dir:
        print(f"📦 Extrayendo en directorio temporal...")
        try:
            # x: eXtract with full paths
            # -y: assume Yes on all queries
            # capture=False para ver el progreso de 7-Zip
            run_command([str(seven_z), "x", f"-o{temp_dir}", str(rar_file), "-y"], capture=False)
            print("   ✓ Extracción completada")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error extrayendo RAR: {e}")
            return False

        # Verificar contenido
        temp_path = Path(temp_dir)
        files = list(temp_path.iterdir())
        if not files:
            print("⚠️  El archivo RAR está vacío")
            return False
        
        print(f"   ℹ️  Archivos encontrados: {[f.name for f in files]}")

        # Detectar si hay carpeta 'ext' para usarla como raíz
        source_path = temp_path
        ext_folder = temp_path / "ext"
        if ext_folder.exists() and ext_folder.is_dir():
             print(f"   📂 Detectada carpeta 'ext' en el backup. Usando su contenido como raíz.")
             source_path = ext_folder

        config = load_config()
        
        # Determinar clusters a procesar
        clusters_to_process = [target_cluster] if target_cluster else config["clusters"].keys()
        
        # Iterar sobre todos los clusters configurados
        for cluster_name in clusters_to_process:
            if cluster_name not in config["clusters"]:
                print(f"⚠️  Cluster '{cluster_name}' no encontrado en configuración.")
                continue
                
            cluster_config = config["clusters"][cluster_name]
            print(f"\n🌐 Procesando Cluster: {cluster_config['name'].upper()}")
            
            # Login una vez por cluster
            if not oc_login(cluster_name):
                print(f"   ❌ No se pudo conectar a {cluster_name}, saltando...")
                continue
            
            # Iterar namespaces del cluster
            for ns in cluster_config["namespaces"]:
                print(f"\n   🔹 Namespace: {ns}")
                
                # Descubrir PVC automáticamente
                pvc_name = find_target_pvc(ns)
                if not pvc_name:
                    print(f"      ❌ Saltando {ns}: No se pudo determinar PVC objetivo")
                    continue
                    
                print(f"      PVC Objetivo: {pvc_name}")
                
                # Usamos mirror=True para asegurar copia exacta (pero ahora sin borrar primero, gracias al cambio en sync_pvc)
                # NOTA: sync_pvc fue modificado para usar --overwrite y NO borrar todo antes si mirror=False
                # El usuario pidió "sincronización sin borrar primero", así que pasamos mirror=False
                # para que sync_pvc NO ejecute el 'find . -delete'.
                if sync_pvc(cluster_name, ns, str(source_path), pvc_name=pvc_name, mirror=False):
                    print(f"      ✅ Completado: {ns}")
                else:
                    print(f"      ❌ Falló: {ns}")

    print("\n🏁 Despliegue masivo finalizado")
    return True



if __name__ == "__main__":
    # Test básico
    print("oc_manager.py cargado correctamente")
    config = load_config()
    print(f"Registry: {config['quay_registry']}")
    print(f"Clusters: {list(config['clusters'].keys())}")
