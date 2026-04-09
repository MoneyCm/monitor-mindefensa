import pandas as pd
from pathlib import Path
import yaml
from logger import log

class DataProcessor:
    """ETL dinámico para archivos Excel de MinDefensa."""
    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as f:
            self.cfg = yaml.safe_load(f)
        self.output_dir = Path("mindefensa_xlsx")
        self.output_dir.mkdir(exist_ok=True)

    def _detectar_cabecera(self, ruta):
        """Busca en qué fila comienzan realmente los datos basándose en alias de columnas."""
        try:
            df_preview = pd.read_excel(ruta, nrows=15, header=None)
            alias_muni = [a.upper() for a in self.cfg['alias_columnas']['municipio']]
            
            for i, row in df_preview.iterrows():
                row_vals = [str(v).upper().strip() for v in row.dropna()]
                if any(alias in row_vals for alias in alias_muni):
                    return i
            return 0
        except Exception as e:
            log.warning(f"No se pudo detectar cabecera en {ruta.name}: {e}")
            return 0

    def procesar_archivo(self, ruta):
        """Lee el Excel, filtra por Jamundí y normaliza columnas."""
        try:
            skip = self._detectar_cabecera(ruta)
            df = pd.read_excel(ruta, skiprows=skip)
            df.columns = [str(c).upper().strip() for c in df.columns]
            
            # Identificar columnas críticas por alias
            alias_muni = self.cfg['alias_columnas']['municipio']
            alias_fecha = self.cfg['alias_columnas']['fecha']
            alias_cant = self.cfg['alias_columnas']['cantidad']
            
            col_muni = next((c for c in df.columns if c in alias_muni), None)
            col_fecha = next((c for c in df.columns if c in alias_fecha), None)
            col_cant = next((c for c in df.columns if c in alias_cant), None)
            
            if not col_muni:
                return {"error": f"Columna de municipio no encontrada (Alias buscados: {alias_muni})"}

            # Filtrado Jamundí
            cod_jamundi = str(self.cfg['municipio']['codigo'])
            # Filtro por código o por nombre (insensible a tildes/case)
            mask = (df[col_muni].astype(str).str.strip() == cod_jamundi) | \
                   (df[col_muni].astype(str).str.upper().str.contains("JAMUNDI", na=False))
            
            df_j = df[mask].copy()
            
            # Normalización de tiempos
            if col_fecha:
                df_j['FECHA_DT'] = pd.to_datetime(df_j[col_fecha], errors='coerce')
                df_j['ANIO'] = df_j['FECHA_DT'].dt.year
                df_j['MES'] = df_j['FECHA_DT'].dt.month
                # Fallback si ANIO es nulo (algunos Excel tienen el año como string en la columna fecha)
                if df_j['ANIO'].isnull().any():
                    df_j['ANIO'] = pd.to_numeric(df_j[col_fecha].astype(str).str[:4], errors='coerce')
            
            # Normalización de totales
            if col_cant:
                raw_cant = df_j[col_cant].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
                df_j['VALOR_NORMALIZADO'] = pd.to_numeric(raw_cant, errors='coerce').fillna(0)
            else:
                df_j['VALOR_NORMALIZADO'] = 1 # Conteo de filas si no hay columna de cantidad (ej. Masacres)

            return {
                "total_nacional": len(df),
                "total_jamundi": int(df_j['VALOR_NORMALIZADO'].sum()) if 'VALOR_NORMALIZADO' in df_j else len(df_j),
                "conteo_jamundi": len(df_j),
                "data": df_j,
                "columnas_detectadas": {
                    "municipio": col_muni,
                    "fecha": col_fecha,
                    "cantidad": col_cant
                }
            }
        except Exception as e:
            return {"error": str(e)}
