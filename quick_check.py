import pandas as pd
from pathlib import Path

xlsx = Path("mindefensa_xlsx/HOMICIDIO INTENCIONAL.xlsx")
JAMUNDI_CODE = 76364

try:
    df = pd.read_excel(xlsx, engine="openpyxl")
    df.columns = [str(c).lower().strip() for c in df.columns]
    col_muni = next((c for c in df.columns if any(x in c for x in ["cod_muni","municipio","mpio"])), None)
    if col_muni:
        mask = (df[col_muni].astype(str).str.strip() == str(JAMUNDI_CODE)) | \
               (df[col_muni].astype(str).str.upper().str.contains("JAMUNDI", na=False))
        df_j = df[mask]
        print(f"HOMICIDIOS JAMUNDI: {len(df_j)}")
        if len(df_j) > 0:
            col_fecha = next((c for c in df.columns if any(x in c for x in ["anio","ano","year","fecha"])), None)
            if col_fecha:
                print(f"Años: {sorted(df_j[col_fecha].dropna().astype(str).str[:4].unique().tolist())}")
    else:
        print("No se encontró columna de municipio")
except Exception as e:
    print(f"Error: {e}")
