"""Send compact, verifiable MinDefensa references to SISC."""

from __future__ import annotations

from datetime import date
import json
import os
from pathlib import Path
import re
from typing import Optional
import unicodedata
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd
import requests

from logger import log


class ReferenceExporter:
    """Keep raw national workbooks in the monitor and upload compact evidence."""

    PRIORITY_DATASETS = (
        "HOMICIDIO INTENCIONAL",
        "HURTO PERSONAS",
        "HURTO DE VEH",
        "EXTORSI",
        "VIOLENCIA INTRAFAMILIAR",
        "LESIONES COMUNES",
    )
    REFERENCE_DEPARTMENTS = {"76", "19"}  # Valle del Cauca y Cauca

    def __init__(self) -> None:
        self.api_url = os.getenv("SISC_API_URL", "https://sisc-backend.onrender.com/api").strip()
        self.token = os.getenv("SISC_TOKEN", "").strip()

    @staticmethod
    def _normalize(value: object) -> str:
        return "_".join(
            unicodedata.normalize("NFD", str(value or ""))
            .encode("ascii", "ignore")
            .decode("ascii")
            .upper()
            .replace("-", " ")
            .split()
        )

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
        if base.endswith("/api"):
            return f"{base}/intelligence/reference-aggregate-upload"
        if "/api/" in base:
            return f"{base.split('/api/', 1)[0]}/api/intelligence/reference-aggregate-upload"
        return f"{base}/api/intelligence/reference-aggregate-upload"

    def _find_header(self, path: Path) -> int:
        preview = pd.read_excel(path, header=None, nrows=20)
        for index, row in preview.iterrows():
            values = {self._normalize(value) for value in row.values if pd.notna(value)}
            if any("COD" in value and "MUNI" in value for value in values) or "MUNICIPIO" in values:
                return index
        raise ValueError("No se encontro la fila de encabezados del archivo de referencia.")

    @staticmethod
    def _column(columns: list[str], candidates: tuple[str, ...]) -> Optional[str]:
        return next((column for column in columns if any(candidate in column for candidate in candidates)), None)

    def _build_payload(self, path: Path, source_cutoff: Optional[str]) -> dict:
        header = self._find_header(path)
        frame = pd.read_excel(path, header=header)
        frame.columns = [self._normalize(column) for column in frame.columns]
        columns = frame.columns.tolist()
        code_column = self._column(columns, ("COD_MUNI", "CVE_MUNI", "CODIGO_MUNICIPIO"))
        municipality_column = self._column(columns, ("MUNICIPIO", "MPIO", "LUGAR"))
        department_column = self._column(columns, ("DEPARTAMENTO", "DEPTO", "DTO"))
        date_column = self._column(columns, ("FECHA_HECHO", "FECHA", "ANIO", "ANO"))
        quantity_column = self._column(columns, ("CANTIDAD", "VICTIMAS", "TOTAL", "NUMERO_CASOS"))
        if not code_column or not date_column:
            raise ValueError("El archivo no tiene codigo DANE o fecha para construir el comparador.")

        codes = (
            frame[code_column].astype(str).str.replace(r"\.0$", "", regex=True)
            .str.replace(r"\D", "", regex=True).str.zfill(5)
        )
        codes = codes.where(codes.str.fullmatch(r"\d{5}"))
        if date_column in {"ANIO", "ANO"}:
            years = pd.to_numeric(frame[date_column], errors="coerce")
            months = pd.Series(1, index=frame.index)
        else:
            dates = pd.to_datetime(frame[date_column], errors="coerce", dayfirst=True)
            years = dates.dt.year
            months = dates.dt.month
        quantities = (
            pd.to_numeric(frame[quantity_column], errors="coerce").fillna(1)
            if quantity_column else pd.Series(1, index=frame.index)
        )
        compact = pd.DataFrame({
            "codigo_dane": codes,
            "municipio": frame[municipality_column].astype(str).str.strip() if municipality_column else codes,
            "departamento": frame[department_column].astype(str).str.strip() if department_column else "NO INFORMADO",
            "anio": years,
            "mes": months,
            "cantidad": quantities,
        })
        minimum_year = date.today().year - 2
        compact = compact.dropna(subset=["codigo_dane", "anio", "mes"])
        compact = compact[(compact["anio"] >= minimum_year) & compact["mes"].between(1, 12)]
        if compact.empty:
            raise ValueError("El archivo no contiene periodos comparables recientes.")
        compact["anio"] = compact["anio"].astype(int)
        compact["mes"] = compact["mes"].astype(int)
        compact["cantidad"] = compact["cantidad"].clip(lower=0).round().astype(int)

        coverage = [
            {"anio": int(year), "municipality_codes": sorted(group["codigo_dane"].unique().tolist())}
            for year, group in compact.groupby("anio")
        ]
        national = compact.groupby(["anio", "mes"], as_index=False)["cantidad"].sum()
        national["codigo_dane"] = "NACIONAL"
        national["municipio"] = "Colombia"
        national["departamento"] = "Colombia"
        regional = compact[compact["codigo_dane"].str[:2].isin(self.REFERENCE_DEPARTMENTS)]
        regional = regional.groupby(
            ["codigo_dane", "municipio", "departamento", "anio", "mes"],
            as_index=False,
        )["cantidad"].sum()
        records = pd.concat([national, regional], ignore_index=True)
        return {
            "filename": path.name,
            "tipo_delito": self._crime_type(path.name),
            "source_cutoff": (source_cutoff or date.today().isoformat())[:10],
            "coverage": coverage,
            "records": records.to_dict(orient="records"),
        }

    def _crime_type(self, filename: str) -> str:
        name = self._normalize(filename)
        if "HOMICIDIO_INTENCIONAL" in name:
            return "Homicidio Intencional"
        if "HURTO_PERSONAS" in name:
            return "Hurto Personas"
        if "HURTO_DE_VEH" in name:
            return "Hurto Vehiculos"
        if "EXTORSI" in name:
            return "Extorsion"
        if "VIOLENCIA_INTRAFAMILIAR" in name:
            return "Violencia Intrafamiliar"
        if "LESIONES_COMUNES" in name:
            return "Lesiones Personales"
        return "Delito General"

    @classmethod
    def is_priority_file(cls, path: Path) -> bool:
        normalized_name = cls._normalize(path.name)
        return any(cls._normalize(pattern) in normalized_name for pattern in cls.PRIORITY_DATASETS)

    @staticmethod
    def _cutoff_date(value: object) -> str:
        if isinstance(value, dict):
            value = value.get("value")
        if isinstance(value, str) and len(value) >= 10:
            return value[:10]
        return date.today().isoformat()

    def export_file(self, path: Path, source_cutoff: Optional[str] = None) -> bool:
        if not self.is_priority_file(path):
            log.info(f"Referencia territorial no requerida para: {path.name}")
            return False

        endpoint = self._endpoint()
        authorization_token = self.token or self._github_oidc_token()
        if not endpoint or not authorization_token:
            log.warning("Referencia territorial omitida: no hay identidad OIDC ni token SISC.")
            return False
        try:
            payload = self._build_payload(path, self._cutoff_date(source_cutoff))
            response = requests.post(
                endpoint,
                headers={"Authorization": f"Bearer {authorization_token}"},
                json=payload,
                timeout=120,
            )
            if response.status_code in (200, 201):
                result = response.json()
                log.info(
                    f"Referencia SISC {result.get('status')}: "
                    f"{result.get('records', 'sin dato')} agregados, "
                    f"anos {result.get('coverage_years', [])}."
                )
                return True
            log.warning(f"Referencia SISC rechazada ({response.status_code}): {response.text[:300]}")
        except (OSError, ValueError, requests.RequestException) as error:
            log.warning(f"No se pudo cargar referencia territorial: {error}")
        return False
