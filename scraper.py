from playwright.sync_api import sync_playwright
import yaml
from logger import log
from utils import retry

class MinDefensaScraper:
    """Scraper basado en Playwright para interceptar la API OCM de MinDefensa."""
    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as f:
            self.cfg = yaml.safe_load(f)
        self.url = "https://www.mindefensa.gov.co/defensa-y-seguridad/datos-y-cifras/informacion-estadistica"
        self.archivos_detectados = []

    def _on_response(self, response):
        """Callback para interceptar respuestas JSON de la API."""
        if "json" in response.headers.get("content-type", "") and response.status == 200:
            try:
                data = response.json()
                self._recorrer_ocm_json(data)
            except:
                pass

    def _recorrer_ocm_json(self, obj):
        """Recursión para encontrar objetos DocumentFile en el JSON de Oracle CM."""
        if isinstance(obj, dict):
            if obj.get("type") == "DocumentFile":
                fields = obj.get("fields", {})
                nombre = (fields.get("name") or fields.get("displayName") or obj.get("name") or "").strip()
                if nombre.upper().endswith(".XLSX"):
                    item_id = obj.get("id")
                    self.archivos_detectados.append({
                        "nombre": nombre,
                        "id": item_id,
                        "fecha": fields.get("updatedDate") or obj.get("updatedDate"),
                        "url": f"https://www.mindefensa.gov.co/sites/web/content/published/api/v1.1/assets/{item_id}/native?siteId=Sitio-Web-Ministerio-Defensa&channelToken=86fd5ad8af1b4db2b56bfc60a05ec867"
                    })
            for v in obj.values():
                if isinstance(v, (dict, list)):
                    self._recorrer_ocm_json(v)
        elif isinstance(obj, list):
            for i in obj:
                self._recorrer_ocm_json(i)

    @retry(Exception, total_tries=2)
    def ejecutar(self):
        """Lanza el navegador y captura la lista de archivos."""
        log.info(f"Iniciando captura en {self.url}")
        self.archivos_detectados = []
        
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True, 
                args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
            )
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
            page = context.new_page()
            
            # Registrar interceptor
            page.on("response", self._on_response)
            
            try:
                page.goto(self.url, wait_until="networkidle", timeout=self.cfg['umbrales']['timeout_playwright'])
            except Exception as e:
                log.warning(f"Timeout o carga parcial en goto: {e}. Intentando procesar lo capturado.")

            # Scroll para disparar lazy loading
            for i in range(6):
                page.evaluate(f"window.scrollBy(0, 700)")
                page.wait_for_timeout(1200)
            
            browser.close()

        # Deduplicar por nombre (clave única)
        unicos = {}
        for a in self.archivos_detectados:
            unicos[a['nombre'].upper().strip()] = a
            
        final_list = list(unicos.values())
        log.info(f"Detección finalizada: {len(final_list)} archivos encontrados.")
        return final_list
