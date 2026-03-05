from playwright.sync_api import sync_playwright
import json, re, pandas as pd, requests, unicodedata
from datetime import datetime
from pathlib import Path

print("="*80)
print("🔍 MONITOR MINDEFENSA - DESCARGADOR AUTOMATICO")
print("="*80)
print(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
print("="*80 + "\n")

URL = "https://www.mindefensa.gov.co/defensa-y-seguridad/datos-y-cifras/informacion-estadistica"
CHANNEL_TOKEN = "86fd5ad8af1b4db2b56bfc60a05ec867"
SITE_ID = "Sitio-Web-Ministerio-Defensa"

ARCHIVOS_INTERES = [
    "HOMICIDIO INTENCIONAL", "LESIONES COMUNES", "VIOLENCIA INTRAFAMILIAR",
    "DELITOS SEXUALES", "SECUESTRO", "EXTORSIÓN", "TERRORISMO", "MASACRES",
    "AFECTACIÓN A LA FUERZA PÚBLICA", "PIRATERÍA TERRESTRE",
    "TRATA DE PERSONAS Y TRÁFICO DE MIGRANTES", "INVASIÓN DE TIERRAS",
    "HURTO PERSONAS", "HURTO A RESIDENCIAS", "HURTO DE VEHÍCULOS",
    "HURTO A COMERCIO", "INCAUTACIÓN DE COCAINA", "INCAUTACIÓN DE MARIHUANA"
]

def norm(t): return ''.join(c for c in unicodedata.normalize('NFD', t.upper()) if unicodedata.category(c) != 'Mn')

archivos_raw = []

def interceptar_respuesta(response):
    if response.status == 200 and "json" in response.headers.get("content-type", "").lower():
        try:
            recorrer_json(response.json())
        except: pass

def recorrer_json(obj, nivel=0):
    if nivel > 8: return
    if isinstance(obj, dict):
        if obj.get("type") == "DocumentFile":
            fields = obj.get("fields", {})
            nombre = (fields.get("name") or fields.get("displayName") or obj.get("name") or "").strip()
            if nombre.upper().endswith(".XLSX"):
                item_id = obj.get("id", "")
                url = f"https://www.mindefensa.gov.co/sites/web/content/published/api/v1.1/assets/{item_id}/native?siteId={SITE_ID}&channelToken={CHANNEL_TOKEN}"
                archivos_raw.append({"nombre": nombre, "url": url})
        for v in obj.values():
            if isinstance(v, (dict, list)): recorrer_json(v, nivel+1)
    elif isinstance(obj, list):
        for item in obj: recorrer_json(item, nivel+1)

with sync_playwright() as p:
    print("🚀 Iniciando navegador...")
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.on("response", interceptar_respuesta)
    print(f"📡 Navegando a: {URL}")
    try:
        page.goto(URL, wait_until='networkidle', timeout=60000)
        page.wait_for_timeout(5000)
        for i in range(5):
            page.evaluate(f"window.scrollTo(0, {i*800})")
            page.wait_for_timeout(700)
    except: pass
    browser.close()

vistos = set()
para_descargar = []
interes_norm = [norm(x) for x in ARCHIVOS_INTERES]
for a in archivos_raw:
    clave = a["nombre"].upper().strip()
    if clave not in vistos:
        vistos.add(clave)
        if any(x in norm(clave) for x in interes_norm):
            para_descargar.append(a)

print(f"Detectados {len(para_descargar)} archivos de interes.")
for item in para_descargar:
    print(f"  -> {item['nombre']}...", end="", flush=True)
    try:
        r = requests.get(item['url'], timeout=90)
        if r.status_code == 200 and len(r.content) > 10000:
            Path(item['nombre']).write_bytes(r.content)
            print(" OK")
        else: print(f" FALLO (Status {r.status_code})")
    except Exception as e: print(f" ERROR: {e}")

print("\n" + "="*80 + "\n✅ PROCESO COMPLETADO\n" + "="*80)
