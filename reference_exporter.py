"""Upload nationwide MinDefensa aggregates without changing the local bulletin."""

from __future__ import annotations

import os
import json
from pathlib import Path
from typing import Optional
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import requests

from logger import log


class ReferenceExporter:
    PRIORITY_DATASETS = (
        "HOMICIDIO INTENCIONAL",
        "HURTO PERSONAS",
        "HURTO DE VEH",
        "EXTORSI",
        "VIOLENCIA INTRAFAMILIAR",
        "LESIONES COMUNES",
    )

    def __init__(self) -> None:
        self.api_url = os.getenv("SISC_API_URL", "https://sisc-backend.onrender.com/api").strip()
        self.token = os.getenv("SISC_TOKEN", "").strip()

    @staticmethod
    def _github_oidc_token() -> Optional[str]:
        request_url = os.getenv("ACTIONS_ID_TOKEN_REQUEST_URL", "").strip()
        request_token = os.getenv("ACTIONS_ID_TOKEN_REQUEST_TOKEN", "").strip()
        if not request_url or not request_token:
            return None
        separator = "&" if "?" in request_url else "?"
        request = Request(
            f"{request_url}{separator}{urlencode({'audience': 'sisc-source-center'})}",
            headers={"Authorization": f"Bearer {request_token}"},
        )
        try:
            with urlopen(request, timeout=10) as response:
                payload = json.loads(response.read().decode("utf-8"))
            return str(payload.get("value")) if payload.get("value") else None
        except (OSError, ValueError) as error:
            log.warning(f"No se pudo solicitar la identidad OIDC: {error}")
            return None

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
        normalized_name = path.name.upper()
        if not any(pattern in normalized_name for pattern in self.PRIORITY_DATASETS):
            log.info(f"Referencia territorial no requerida para: {path.name}")
            return False

        endpoint = self._endpoint()
        authorization_token = self.token or self._github_oidc_token()
        if not endpoint or not authorization_token:
            log.warning("Referencia territorial omitida: no hay identidad OIDC ni token SISC.")
            return False

        data = {"source_cutoff": source_cutoff} if source_cutoff else {}
        headers = {"Authorization": f"Bearer {authorization_token}"}
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
