import pandas as pd
import requests
import os
from pathlib import Path

# Cargar el listado detectado por la API
csv_path = 'listado_mindefensa_api.csv'
if not os.path.exists(csv_path):
    print(f"Error: {csv_path} no existe.")
    exit(1)

df = pd.read_csv(csv_path)

# Delitos que el boletin espera (DATASETS original)
ARCHIVOS_INTERES = [
    "HOMICIDIO INTENCIONAL", "LESIONES COMUNES", "VIOLENCIA INTRAFAMILIAR",
    "DELITOS SEXUALES", "SECUESTRO", "EXTORSIÓN", "TERRORISMO", "MASACRES",
    "AFECTACIÓN A LA FUERZA PÚBLICA", "PIRATERÍA TERRESTRE",
    "TRATA DE PERSONAS Y TRÁFICO DE MIGRANTES", "INVASIÓN DE TIERRAS",
    "HURTO PERSONAS", "HURTO A RESIDENCIAS", "HURTO DE VEHÍCULOS",
    "HURTO A COMERCIO", "INCAUTACIÓN DE COCAINA", "INCAUTACIÓN DE MARIHUANA"
]

print(f"Descargando {len(ARCHIVOS_INTERES)} archivos de interes...")

for _, row in df.iterrows():
    nombre = str(row['nombre']).strip()
    url = str(row['url_descarga']).strip()
    
    if nombre.upper().endswith('.XLSX'):
        import unicodedata
        def norm(t): return ''.join(c for c in unicodedata.normalize('NFD', t.upper()) if unicodedata.category(c) != 'Mn')
        
        nombre_norm = norm(nombre)
        interes_norm = [norm(x) for x in ARCHIVOS_INTERES]
        
        if any(x in nombre_norm for x in interes_norm):
            print(f"  -> {nombre}...", end="", flush=True)
            try:
                r = requests.get(url, timeout=90)
                if r.status_code == 200 and len(r.content) > 5000:
                    Path(nombre).write_bytes(r.content)
                    print(" OK")
                else:
                    print(f" FALLO (HTTP {r.status_code} o archivo pequeño)")
            except Exception as e:
                print(f" ERROR: {e}")

print("\nDescargas finalizadas.")
