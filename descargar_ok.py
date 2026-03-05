import requests
import json
import os
from pathlib import Path

# Configuracion capturada
SITE_ID = "Sitio-Web-Ministerio-Defensa"
# El token debe ser el mas reciente (obtenido del script anterior)
TOKEN = "86fd5ad8af1b4db2b56bfc60a05ec867"

DELITOS = {
    "HOMICIDIO INTENCIONAL.xlsx": "CONT0D5438843109405FAD19010C7359F679",
    "LESIONES COMUNES.xlsx": "CONT53D83D55A5514BC2B49CD3940F8EF90B",
    "VIOLENCIA INTRAFAMILIAR.xlsx": "CONT9A98C7E7795349B89694F2B02B5E26E5",
    "DELITOS SEXUALES.xlsx": "CONTB8622EB0B12F47239C29323F6783B097",
    "SECUESTRO.xlsx": "CONTAEA231662973499480EDF6B086846152",
    "EXTORSIÓN.xlsx": "CONT01A9908F873541E4AF697D63F4CB6ED4",
    "TERRORISMO.xlsx": "CONT252D3F95146C49A296D4686A39D9E0BC",
    "MASACRES.xlsx": "CONT7367B80970344917B3C39535B5811F7E",
    "AFECTACIÓN A LA FUERZA PÚBLICA.xlsx": "CONT7253457A0266453086E5007559B283AF",
    "PIRATERÍA TERRESTRE.xlsx": "CONT2588106D9410484BA4A9684128D4D652",
    "TRATA DE PERSONAS Y TRÁFICO DE MIGRANTES.xlsx": "CONT72DB52E106D14D1A8A2F2B2D19890F3B",
    "INVASIÓN DE TIERRAS.xlsx": "CONT5B0E131E7FA3430B82D7040AF33F6E7B",
    "HURTO PERSONAS.xlsx": "CONT492215DBB0E14D88916327668636E345",
    "HURTO A RESIDENCIAS.xlsx": "CONTECC8648B77CD41E59C731A88A96C247A",
    "HURTO DE VEHÍCULOS.xlsx": "CONT93922D97D87E4CD9A35F48651CB135C2",
    "HURTO A COMERCIO.xlsx": "CONT8788DA2A7DB44D8787C1743E51F72960",
    "INCAUTACIÓN DE COCAINA.xlsx": "CONTBCC6014457314272BAE7C74A256E6686",
    "INCAUTACIÓN DE MARIHUANA.xlsx": "CONT72DB3142DB5A4C37BAE03D811409B1BB"
}

def descargar():
    for nombre, item_id in DELITOS.items():
        url = f"https://www.mindefensa.gov.co/sites/web/content/published/api/v1.1/assets/{item_id}/native?siteId={SITE_ID}&channelToken={TOKEN}"
        print(f"Bajando {nombre}...", end="", flush=True)
        try:
            r = requests.get(url, timeout=60)
            if r.status_code == 200 and len(r.content) > 10000:
                Path(nombre).write_bytes(r.content)
                print(" OK")
            else:
                print(f" ERROR (Status {r.status_code}, Size {len(r.content)})")
        except Exception as e:
            print(f" ERROR: {e}")

if __name__ == "__main__":
    descargar()
