import json
from pathlib import Path
from logger import log
from utils import get_date_value

class StateManager:
    """Gestiona la persistencia del estado en mindefensa_state.json."""
    def __init__(self, state_file="mindefensa_state.json"):
        self.path = Path(state_file)
        self.data = self.load()

    def load(self):
        if not self.path.exists():
            return {"archivos": {}, "ultima_revision": None}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                return json.loads(content) if content else {"archivos": {}, "ultima_revision": None}
        except Exception as e:
            log.error(f"Error cargando estado: {e}. Iniciando estado vacío.")
            return {"archivos": {}, "ultima_revision": None}

    def save(self, unicos_detectados, nuevos_count, cambios_count):
        self.data["ultima_revision"] = Path("date_ref.txt").stat().st_mtime if Path("date_ref.txt").exists() else None # Placeholder
        from datetime import datetime
        self.data["ultima_revision"] = datetime.now().isoformat()
        self.data["nuevos_ultimo"] = nuevos_count
        self.data["cambios_ultimo"] = cambios_count
        self.data["archivos"] = {a["nombre"]: {"fecha": a["fecha"], "id": a["id"]} for a in unicos_detectados}
        
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
            log.info(f"Estado guardado en {self.path}")
        except Exception as e:
            log.error(f"No se pudo guardar el archivo de estado: {e}")

    def comparar(self, unicos_encontrados):
        nuevos, cambiados = [], []
        previos = self.data.get("archivos", {})
        
        # Normalizar para comparación insensible a mayúsculas
        previos_norm = {k.upper().strip(): v for k, v in previos.items()}
        
        for a in unicos_encontrados:
            nombre = a["nombre"]
            clave = nombre.upper().strip()
            
            if clave not in previos_norm:
                nuevos.append(a)
            else:
                fecha_act = get_date_value(a["fecha"])
                fecha_prev = get_date_value(previos_norm[clave].get("fecha"))
                if fecha_act != fecha_prev:
                    cambiados.append(a)
        
        return nuevos, cambiados
