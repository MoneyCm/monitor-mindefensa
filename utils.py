import time
import hashlib
from functools import wraps
from logger import log

def retry(exceptions, total_tries=3, initial_wait=2, backoff=2):
    """Decorador para reintentar funciones con backoff exponencial."""
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            tries, delay = total_tries, initial_wait
            while tries > 1:
                try:
                    return f(*args, **kwargs)
                except exceptions as e:
                    log.warning(f"Falla en {f.__name__}: {e}. Reintentando en {delay}s... ({tries-1} restantes)")
                    time.sleep(delay)
                    tries -= 1
                    delay *= backoff
            return f(*args, **kwargs)
        return wrapper
    return decorator

def calculate_sha256(file_path):
    """Calcula el hash SHA256 de un archivo."""
    sha256_hash = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        log.error(f"Error calculando SHA256 para {file_path}: {e}")
        return "N/A"

def get_date_value(f_obj):
    """Normaliza valores de fecha que pueden venir como string o dict (OCM API)."""
    if isinstance(f_obj, dict):
        return f_obj.get("value", "")
    return str(f_obj)
