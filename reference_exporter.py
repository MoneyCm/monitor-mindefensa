"""Upload nationwide MinDefensa aggregates without changing the local bulletin."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

import requests

from logger import log


class ReferenceExporter:
    def __init__(self) -> None:
        self.api_url = os.getenv("SISC_API_URL", "").strip()
        self.token = os.getenv("SISC_TOKEN", "").strip()

    def _endpoint(self) -> Optional[str]:
        if not self.api_url:
            return None
        base = self.api_url.rstrip("/")
        if base.endswith("/intelligence/upload"):
            return f"{base[:-len('/upload')]}/reference-upload"
        if base.endswith("/api"):
            return f"{base}/intelligence/reference-upload"
        if "/api/" in base:
            return f"{base.split('/api/', 1)[0]}/api/intelligence/reference-upload"
        return f"{base}/api/intelligence/reference-upload"

    def export_file(self, path: Path, source_cutoff: Optional[str] = None) -> bool:
        endpoint = self._endpoint()
        if not endpoint or not self.token:
            log.warning("Referencia territorial omitida: falta SISC_API_URL o SISC_TOKEN.")
            return False

        data = {"source_cutoff": source_cutoff} if source_cutoff else {}
        headers = {"Authorization": f"Bearer {self.token}"}
        try:
            with path.open("rb") as stream:
                response = requests.post(
                    endpoint,
                    headers=headers,
                    data=data,
                    files={"file": (path.name, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                    timeout=180,
                )
            if response.status_code in (200, 201):
                payload = response.json()
                log.info(
                    f"Referencia SISC {payload.get('status')}: "
                    f"{payload.get('municipalities', 'sin dato')} municipios, "
                    f"{payload.get('records', 'sin dato')} agregados."
                )
                return True
            log.warning(f"Referencia SISC rechazada ({response.status_code}): {response.text[:300]}")
        except requests.RequestException as error:
            log.warning(f"No se pudo cargar referencia territorial: {error}")
        return False
