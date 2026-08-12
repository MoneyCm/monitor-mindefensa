"""Report the MinDefensa historical monitor status to the SISC source center."""

from __future__ import annotations

import json
import os
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_API_URL = "https://sisc-backend.onrender.com/api"


def _utc_now(now: Optional[datetime] = None) -> datetime:
    value = now or datetime.now(timezone.utc)
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).replace(microsecond=0)


def _iso_datetime(value: datetime) -> str:
    return value.isoformat().replace("+00:00", "Z")


def _load_json(path: Path) -> Optional[Dict[str, Any]]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else None
    except (OSError, json.JSONDecodeError):
        return None


def _parse_date(value: Any) -> Optional[date]:
    if not value:
        return None
    for format_string in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(value).strip(), format_string).date()
        except ValueError:
            continue
    return None


def _as_non_negative_int(value: Any) -> int:
    try:
        return max(int(value), 0)
    except (TypeError, ValueError):
        return 0


def build_payload(
    summary_path: Path = BASE_DIR / "resumen_actual.json",
    state_path: Path = BASE_DIR / "mindefensa_state.json",
    *,
    outcome: str = "success",
    data_changed: bool = False,
    now: Optional[datetime] = None,
) -> Dict[str, Any]:
    checked_at = _utc_now(now)
    normalized_outcome = (outcome or "success").strip().lower()
    workflow_ok = normalized_outcome == "success"
    summary = _load_json(Path(summary_path)) or {}
    state = _load_json(Path(state_path)) or {}
    raw_indicators = summary.get("indicadores")
    indicators = raw_indicators if isinstance(raw_indicators, list) else []
    cutoffs = [
        parsed
        for item in indicators
        if isinstance(item, dict)
        for parsed in [_parse_date(item.get("ultimo_registro"))]
        if parsed is not None
    ]
    cutoff = max(cutoffs) if cutoffs else None
    record_count = sum(
        _as_non_negative_int(item.get("actual"))
        for item in indicators
        if isinstance(item, dict)
    )
    discovered_assets = state.get("archivos") if isinstance(state.get("archivos"), dict) else {}
    warnings = []

    if not workflow_ok:
        warnings.append("La revision diaria de MinDefensa termino con error.")
    elif not summary:
        warnings.append("La fuente fue revisada, pero no existe un resumen historico procesado.")
    elif not cutoff:
        warnings.append("El resumen historico no informo una fecha de corte valida.")
    if workflow_ok and not discovered_assets:
        warnings.append("No se encontro el inventario de archivos de MinDefensa.")

    if not workflow_ok:
        status, quality = "ERROR", "ERROR"
    elif not cutoff:
        status, quality = "NEEDS_REVIEW", "INCOMPLETE"
    elif warnings:
        status, quality = "CURRENT", "WARNING"
    else:
        status, quality = "CURRENT", "VALIDATED"

    payload: Dict[str, Any] = {
        "connector_code": "MINDEFENSA",
        "status": status,
        "quality_status": quality,
        "last_checked_at": _iso_datetime(checked_at),
        "record_count": record_count,
        "indicator_count": len(indicators),
        "warnings": warnings,
        "details": {
            "workflow": os.getenv("GITHUB_WORKFLOW", "monitor-mindefensa"),
            "run_id": os.getenv("GITHUB_RUN_ID"),
            "outcome": normalized_outcome,
            "data_changed": bool(data_changed),
            "discovered_assets": len(discovered_assets),
            "check_mode": "official-content-api-metadata-first",
            "source_role": "historical-backup",
        },
    }
    if cutoff:
        payload["source_cutoff_date"] = cutoff.isoformat()
        payload["period_label"] = (
            f"Corte al {cutoff.isoformat()} - {len(indicators)} indicadores historicos"
        )
    if workflow_ok:
        payload["last_success_at"] = _iso_datetime(checked_at)
    if workflow_ok and data_changed:
        payload["last_change_detected_at"] = _iso_datetime(checked_at)
    return payload


def _heartbeat_url(api_url: str) -> str:
    base = api_url.strip().rstrip("/")
    return base if base.endswith("/source-center/heartbeat") else f"{base}/source-center/heartbeat"


def _request_github_oidc_token(audience: str = "sisc-source-center") -> Optional[str]:
    request_url = os.getenv("ACTIONS_ID_TOKEN_REQUEST_URL", "").strip()
    request_token = os.getenv("ACTIONS_ID_TOKEN_REQUEST_TOKEN", "").strip()
    if not request_url or not request_token:
        return None
    separator = "&" if "?" in request_url else "?"
    request = Request(
        f"{request_url}{separator}{urlencode({'audience': audience})}",
        headers={"Authorization": f"Bearer {request_token}"},
    )
    try:
        with urlopen(request, timeout=10) as response:
            result = json.loads(response.read().decode("utf-8"))
        return str(result.get("value")) if isinstance(result, dict) and result.get("value") else None
    except (HTTPError, URLError, TimeoutError, OSError, json.JSONDecodeError) as error:
        print(f"[AVISO] No se pudo obtener la identidad OIDC de GitHub: {error}.")
        return None


def send_heartbeat(
    payload: Dict[str, Any],
    *,
    api_url: Optional[str] = None,
    service_key: Optional[str] = None,
    oidc_token: Optional[str] = None,
    timeout: int = 20,
) -> bool:
    token = _request_github_oidc_token() if oidc_token is None else oidc_token.strip()
    key = (service_key if service_key is not None else os.getenv("SISC_SOURCE_MONITOR_KEY", "")).strip()
    if not token and not key:
        print("[AVISO] Heartbeat SISC omitido: no hay identidad OIDC ni clave de servicio.")
        return False

    headers = {"Content-Type": "application/json", "User-Agent": "monitor-mindefensa/1.0"}
    headers["Authorization" if token else "X-SISC-SOURCE-KEY"] = f"Bearer {token}" if token else key
    request = Request(
        _heartbeat_url(api_url or os.getenv("SISC_API_URL", DEFAULT_API_URL)),
        data=json.dumps(payload, ensure_ascii=True).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            accepted = 200 <= response.status < 300
        print(f"[INFO] Heartbeat SISC enviado ({payload['status']}).")
        return accepted
    except HTTPError as error:
        print(f"[AVISO] El API SISC rechazo el heartbeat (HTTP {error.code}).")
    except (URLError, TimeoutError, OSError) as error:
        print(f"[AVISO] No se pudo enviar el heartbeat SISC: {error}.")
    return False


def main() -> int:
    changed = os.getenv("DATA_CHANGED", "false").strip().lower() == "true"
    payload = build_payload(
        outcome=os.getenv("SISC_MONITOR_OUTCOME", "success"),
        data_changed=changed,
    )
    print(
        "[INFO] Estado para Centro de fuentes: "
        f"{payload['status']} / {payload['quality_status']} / "
        f"{payload['indicator_count']} indicadores."
    )
    return 0 if send_heartbeat(payload) else 1


if __name__ == "__main__":
    raise SystemExit(main())
