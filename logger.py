import sys
from loguru import logger

def setup_app_logger():
    """Configura Loguru para salida estándar y archivo JSON."""
    logger.remove()
    # Log en consola (limpio para GitHub Actions)
    logger.add(sys.stderr, format="<green>{time:HH:mm:ss}</green> | <level>{level: <7}</level> | <cyan>{function}</cyan> - <level>{message}</level>")
    # Log JSON para debugging profundo
    logger.add("monitor_execution.json", serialize=True, rotation="1 day")
    return logger

log = setup_app_logger()
