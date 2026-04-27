"""
Módulo de almacenamiento: config, autenticación, encriptación y stats.
"""
import json, os, sys, hashlib, base64, socket

try:
    from cryptography.fernet import Fernet
    CRYPTO_OK = True
except ImportError:
    CRYPTO_OK = False


# ============================================================
#  RUTAS
# ============================================================

def get_app_dir():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_data_dir():
    """Directorio persistente del usuario."""
    base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")),
                        "HMConsulting", "CargadorHoras")
    os.makedirs(base, exist_ok=True)
    return base


def get_asset(name):
    """Busca un asset (logo, ícono, template) en varias ubicaciones."""
    # En modo congelado (PyInstaller)
    if getattr(sys, "frozen", False):
        path = os.path.join(sys._MEIPASS, name)
        if os.path.exists(path): return path
    # En modo desarrollo: junto al script principal
    path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), name)
    if os.path.exists(path): return path
    return name


# Archivos de datos
CONFIG_FILE   = os.path.join(get_data_dir(), "config.json")
PASS_FILE     = os.path.join(get_data_dir(), "auth.json")
CLIENTES_FILE = os.path.join(get_data_dir(), "clientes.json")
STATS_FILE    = os.path.join(get_data_dir(), "stats.json")


# ============================================================
#  AUTENTICACIÓN
# ============================================================

def cargar_auth():
    if os.path.exists(PASS_FILE):
        try:
            return json.load(open(PASS_FILE, encoding="utf-8"))
        except Exception as e:
            print(f"[auth] Error leyendo auth.json: {e}")
    return {"hash": None}


def guardar_auth(auth):
    json.dump(auth, open(PASS_FILE, "w", encoding="utf-8"), indent=2)


def hash_pass(password):
    return hashlib.sha256(password.encode()).hexdigest()


# ============================================================
#  ENCRIPTACIÓN
# ============================================================

def get_fernet(password=None):
    """Devuelve un Fernet con clave combinada de contraseña + máquina."""
    if not CRYPTO_OK:
        return None
    machine_raw = hashlib.sha256(
        f"{os.environ.get('USERNAME','user')}@{socket.gethostname()}".encode()
    ).digest()
    if password:
        pass_raw = hashlib.sha256(password.encode()).digest()
        combined = bytes(a ^ b for a, b in zip(machine_raw, pass_raw))
    else:
        combined = machine_raw
    key = base64.urlsafe_b64encode(combined)
    return Fernet(key)


def encriptar(texto, password=None):
    f = get_fernet(password)
    if not f: return texto
    return f.encrypt(texto.encode()).decode()


def desencriptar(texto_enc, password=None):
    f = get_fernet(password)
    if not f: return texto_enc
    try:
        return f.decrypt(texto_enc.encode()).decode()
    except Exception:
        return ""


# ============================================================
#  CONSTANTES DE NEGOCIO
# ============================================================

ACTIVIDADES = [
    "Soporte", "Análisis", "Configuración", "Desarrollo",
    "Diseño", "Documentación", "Capacitación", "Migración de Datos",
    "Pruebas Unitarias", "Pruebas Integrales", "Reunión", "Otra",
]

UBICACIONES = [
    "Oficina Rosario", "Oficina BsAs", "Oficina Venado",
    "Remoto", "On Site",
]

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
TIPOS_HORA  = ["Funcional", "SSFF", "ABAP"]


# ============================================================
#  CONFIGURACIÓN
# ============================================================

CONFIG_DEFAULT = {
    "redmine_url":      "https://rdm.hmconsulting.com.ar",
    "api_key":          "",
    "archivo_excel":    "",
    "hoja":             "Horas",
    "actividad":        "Soporte",
    "tipo_hora":        "Funcional",
    "ubicacion":        "Oficina Rosario",
    "ubicacion_remoto": "Remoto",
    "dia_remoto":       -1,
    "usa_id_ticket":    True,
}


def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            cfg = CONFIG_DEFAULT.copy()
            data = json.load(open(CONFIG_FILE, encoding="utf-8"))
            if data.get("_enc") and CRYPTO_OK:
                try:
                    data["api_key"] = desencriptar(data["api_key"])
                except Exception as e:
                    print(f"[config] No se pudo desencriptar API key: {e}")
                    data["api_key"] = ""
            cfg.update(data)
            return cfg
        except Exception as e:
            print(f"[config] Error leyendo config.json: {e}")
    return CONFIG_DEFAULT.copy()


def guardar_config(cfg, password=None):
    data = dict(cfg)
    if CRYPTO_OK and data.get("api_key"):
        data["api_key"] = encriptar(data["api_key"], password)
        data["_enc"] = True
    else:
        data.pop("_enc", None)
    json.dump(data, open(CONFIG_FILE, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


def cargar_clientes():
    if os.path.exists(CLIENTES_FILE):
        try:
            return json.load(open(CLIENTES_FILE, encoding="utf-8"))
        except Exception as e:
            print(f"[clientes] Error leyendo clientes.json: {e}")
    return []


def guardar_clientes(lista):
    json.dump(lista, open(CLIENTES_FILE, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)


# ============================================================
#  ESTADÍSTICAS
# ============================================================

def cargar_stats():
    if os.path.exists(STATS_FILE):
        try:
            return json.load(open(STATS_FILE, encoding="utf-8"))
        except Exception:
            pass
    return {"horas_cargadas": 0, "issues_creados": 0, "sesiones": 0}


def guardar_stats(s):
    json.dump(s, open(STATS_FILE, "w", encoding="utf-8"), indent=2)
