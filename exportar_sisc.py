import os
import requests
import json
import yaml
from logger import log

class SISCExporter:
    """Módulo para exportar datos procesados a la API del SISC Jamundí Pro."""
    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as f:
            self.cfg = yaml.safe_load(f)
        self.api_url = os.environ.get("SISC_API_URL")
        self.token = os.environ.get("SISC_TOKEN")

    def exportar(self, resultados):
        """Envía los datos consolidados por delito/anio/mes a la API."""
        if not self.api_url or not self.token:
            log.info("API SISC no configurada. Saltando exportación.")
            return

        exitos, fallos = 0, 0
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

        for delito, r in resultados.items():
            if "error" in r or "data" not in r: continue
            
            df = r['data']
            # Consolidar por Año y Mes
            if 'ANIO' in df.columns and 'MES' in df.columns:
                consolidado = df.groupby(['ANIO', 'MES'])['VALOR_NORMALIZADO'].sum().reset_index()
                
                for _, row in consolidado.iterrows():
                    payload = {
                        "municipio": self.cfg['municipio']['nombre'],
                        "codigo_dane": self.cfg['municipio']['codigo'],
                        "delito": delito,
                        "anio": int(row['ANIO']),
                        "mes": int(row['MES']),
                        "valor": float(row['VALOR_NORMALIZADO']),
                        "fuente": self.cfg['municipio']['fuente'],
                        "metadata": r.get('columnas_detectadas', {})
                    }
                    
                    try:
                        resp = requests.post(self.api_url, json=payload, headers=headers, timeout=10)
                        if resp.status_code in (200, 201):
                            exitos += 1
                        else:
                            fallos += 1
                            log.debug(f"Falla exportando {delito} {row['ANIO']}-{row['MES']}: {resp.status_code}")
                    except Exception as e:
                        fallos += 1
                        log.debug(f"Error conexión API: {e}")

        log.info(f"Exportación SISC finalizada. Éxitos: {exitos}, Fallos: {fallos}")
