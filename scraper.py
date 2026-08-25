import os
from typing import Any

import requests
import yaml

from logger import log
from utils import retry


class MinDefensaScraper:
    """Discover published XLSX assets, using the official content API first."""

    PAGE_URL = (
        "https://www.mindefensa.gov.co/defensa-y-seguridad/"
        "datos-y-cifras/informacion-estadistica"
    )
    API_URL = (
        "https://www.mindefensa.gov.co/sites/web/content/published/"
        "api/v1.1/items"
    )
    ASSET_URL = (
        "https://www.mindefensa.gov.co/sites/web/content/published/"
        "api/v1.1/assets/{asset_id}/native"
    )
    SITE_ID = "Sitio-Web-Ministerio-Defensa"
    CHANNEL_TOKEN = "86fd5ad8af1b4db2b56bfc60a05ec867"

    # Only the families used by the Jamundi historical contrast are queried.
    FILE_CATEGORIES = (
        "Estad\u00edstica Desagregada-Delitos contra la vida y la integridad personal",
        "Estad\u00edstica Desagregada-Delitos contra el patrimonio econ\u00f3mico",
        "Estad\u00edstica Desagregada-Delitos contra la familia",
        "Estad\u00edstica Desagregada-Delitos contra la libertad, integridad y formaci\u00f3n sexuales",
        "Estad\u00edstica Desagregada-Delitos contra la libertad individual y otras garant\u00edas constitucionales",
        "Estad\u00edstica Desagregada-Delitos contra la seguridad p\u00fablica",
        "Estad\u00edstica Desagregada-Delitos contra la protecci\u00f3n de la informaci\u00f3n y de los datos",
    )

    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as file:
            self.cfg = yaml.safe_load(file)
        self.archivos_detectados = []

    @classmethod
    def _asset_from_item(cls, item: dict[str, Any]):
        fields = item.get("fields") or {}
        name = (
            item.get("name")
            or fields.get("name")
            or fields.get("displayName")
            or ""
        ).strip()
        asset_id = item.get("id")
        category = str(fields.get("filecategory") or "").strip()
        has_xlsx_extension = name.upper().endswith(".XLSX")

        # MinDefensa republished the statistical assets without filename
        # extensions in August 2026. Category-scoped API results still point to
        # XLSX content, so preserve the stable local filename expected by the
        # downloader, state file and ETL pipeline.
        is_extensionless_statistical_asset = (
            bool(name)
            and "." not in name
            and category in cls.FILE_CATEGORIES
        )
        if not asset_id or not (has_xlsx_extension or is_extensionless_statistical_asset):
            return None
        if is_extensionless_statistical_asset:
            name = f"{name}.xlsx"
        return {
            "nombre": name,
            "id": asset_id,
            "fecha": item.get("updatedDate") or fields.get("updatedDate"),
            "url": cls.ASSET_URL.format(asset_id=asset_id)
            + f"?siteId={cls.SITE_ID}&channelToken={cls.CHANNEL_TOKEN}",
        }

    @staticmethod
    def _deduplicate(assets):
        unique = {}
        for asset in assets:
            unique[asset["nombre"].upper().strip()] = asset
        return list(unique.values())

    def _api_params(self, category, offset):
        return {
            "siteId": self.SITE_ID,
            "limit": 100,
            "offset": offset,
            "orderBy": "name:asc",
            "q": f'type eq "DocumentFile" and fields.filecategory eq "{category}"',
        }

    def _discover_via_api(self):
        session = requests.Session()
        session.headers.update(
            {
                "Accept": "application/json",
                "User-Agent": "SISC-Jamundi-Source-Monitor/1.0",
            }
        )
        assets = []
        for category in self.FILE_CATEGORIES:
            offset = 0
            while True:
                response = session.get(
                    self.API_URL,
                    params=self._api_params(category, offset),
                    timeout=30,
                )
                response.raise_for_status()
                payload = response.json()
                items = payload.get("items") or []
                for item in items:
                    asset = self._asset_from_item(item)
                    if asset:
                        assets.append(asset)
                if not payload.get("hasMore") or not items:
                    break
                offset += len(items)
        return self._deduplicate(assets)

    def _walk_browser_payload(self, value):
        if isinstance(value, dict):
            if value.get("type") == "DocumentFile":
                asset = self._asset_from_item(value)
                if asset:
                    self.archivos_detectados.append(asset)
            for child in value.values():
                if isinstance(child, (dict, list)):
                    self._walk_browser_payload(child)
        elif isinstance(value, list):
            for child in value:
                self._walk_browser_payload(child)

    def _discover_via_browser(self):
        from playwright.sync_api import sync_playwright

        self.archivos_detectados = []

        def on_response(response):
            if response.status != 200:
                return
            if "json" not in response.headers.get("content-type", "").lower():
                return
            try:
                self._walk_browser_payload(response.json())
            except Exception:
                return

        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
            )
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            )
            page = context.new_page()
            page.on("response", on_response)
            try:
                page.goto(
                    self.PAGE_URL,
                    wait_until="networkidle",
                    timeout=self.cfg["umbrales"]["timeout_playwright"],
                )
            except Exception as error:
                log.warning(f"Carga parcial en el respaldo de navegador: {error}")
            for _ in range(6):
                page.evaluate("window.scrollBy(0, 700)")
                page.wait_for_timeout(1200)
            browser.close()
        return self._deduplicate(self.archivos_detectados)

    @retry(Exception, total_tries=2)
    def ejecutar(self):
        log.info("Consultando metadatos de MinDefensa mediante la API oficial.")
        try:
            assets = self._discover_via_api()
            if assets:
                log.info(f"Deteccion finalizada por API: {len(assets)} archivos.")
                return assets
            raise RuntimeError("La API no devolvio archivos XLSX de interes.")
        except Exception as error:
            if os.environ.get("MINDEFENSA_BROWSER_FALLBACK", "false").lower() != "true":
                raise RuntimeError(f"No fue posible consultar la API de MinDefensa: {error}") from error
            log.warning(f"API no disponible; usando respaldo de navegador: {error}")
            assets = self._discover_via_browser()
            if not assets:
                raise RuntimeError("El respaldo de navegador tampoco encontro archivos.")
            log.info(f"Deteccion finalizada por navegador: {len(assets)} archivos.")
            return assets
