"""Alertas operativas del monitor: avisa por correo cuando una corrida falla.

Se ejecuta desde el workflow como paso `if: failure()` y nunca debe tumbar
el job (siempre sale con codigo 0): su unico trabajo es notificar.
"""

from __future__ import annotations

import argparse
import json
import os
import smtplib
import sys
from datetime import datetime, timezone
from email.mime.text import MIMEText
from pathlib import Path

import yaml

from logger import log

BASE_DIR = Path(__file__).resolve().parent


def cargar_config(config_path: Path = BASE_DIR / "config.yaml") -> dict:
    with open(config_path, "r", encoding="utf-8") as file:
        return yaml.safe_load(file) or {}


def destinatarios_alerta(cfg: dict) -> list[str]:
    correo = cfg.get("correo", {}) if isinstance(cfg, dict) else {}
    especificos = correo.get("destinatarios_alertas")
    if especificos:
        return list(especificos)
    return list(correo.get("destinatarios", []))


def dias_sin_importacion(
    state_path: Path = BASE_DIR / "mindefensa_state.json",
    ahora: datetime | None = None,
) -> int | None:
    """Dias desde la ultima revision exitosa. None si no hay estado previo."""
    try:
        state = json.loads(Path(state_path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    ultima = (state or {}).get("ultima_revision")
    if not ultima:
        return None
    try:
        fecha = datetime.fromisoformat(str(ultima).replace("Z", "+00:00"))
    except ValueError:
        return None
    if fecha.tzinfo is None:
        fecha = fecha.replace(tzinfo=timezone.utc)
    referencia = ahora or datetime.now(timezone.utc)
    if referencia.tzinfo is None:
        referencia = referencia.replace(tzinfo=timezone.utc)
    return max((referencia - fecha).days, 0)


def extraer_errores(
    log_path: Path = BASE_DIR / "monitor_execution.json",
    limite: int = 5,
) -> list[str]:
    """Ultimos registros ERROR/CRITICAL del log JSON de loguru."""
    errores: list[str] = []
    try:
        lineas = Path(log_path).read_text(encoding="utf-8").splitlines()
    except OSError:
        return errores
    for linea in reversed(lineas):
        try:
            registro = json.loads(linea)
        except json.JSONDecodeError:
            continue
        nivel = ((registro.get("record") or {}).get("level") or {}).get("name", "")
        if nivel not in ("ERROR", "CRITICAL"):
            continue
        errores.append(str(registro.get("text", "")).strip()[:500])
        if len(errores) >= limite:
            break
    return list(reversed(errores))


def url_run() -> str | None:
    servidor = os.environ.get("GITHUB_SERVER_URL", "https://github.com").rstrip("/")
    repo = os.environ.get("GITHUB_REPOSITORY", "").strip()
    run_id = os.environ.get("GITHUB_RUN_ID", "").strip()
    if repo and run_id:
        return f"{servidor}/{repo}/actions/runs/{run_id}"
    return None


def construir_alerta_fallo(
    dias: int | None,
    errores: list[str],
    umbral_urgente: int = 7,
) -> tuple[str, str]:
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    if dias is None:
        estado = "sin importaciones exitosas registradas"
        urgente = True
    else:
        estado = f"ultima importacion exitosa hace {dias} dia(s)"
        urgente = dias >= umbral_urgente
    prefijo = "[URGENTE]" if urgente else "[AVISO]"
    asunto = f"{prefijo} Monitor MinDefensa fallo ({fecha}) - {estado}"

    lineas = [
        "Cordial saludo,",
        "",
        "La corrida programada del Monitor MinDefensa V2 fallo.",
        f"Estado de la fuente: {estado}.",
    ]
    enlace = url_run()
    if enlace:
        lineas += ["", f"Ver el run: {enlace}"]
    if errores:
        lineas += ["", "Ultimos errores registrados:"]
        lineas += [f"- {error}" for error in errores]
    lineas += [
        "",
        "Si la causa es la caida de mindefensa.gov.co (HTTP 503), no se requiere",
        "accion en el monitor: se retomara solo en la siguiente corrida exitosa.",
        "Si el error indica un cambio de esquema de la API, revisar scraper.py.",
    ]
    return asunto, "\n".join(lineas)


def enviar(asunto: str, cuerpo: str, cfg: dict | None = None) -> bool:
    cfg = cfg if cfg is not None else cargar_config()
    correo_cfg = cfg.get("correo", {}) if isinstance(cfg, dict) else {}
    usuario = os.environ.get("GMAIL_USER", "").strip()
    clave = os.environ.get("GMAIL_PASS", "").strip()
    destinos = destinatarios_alerta(cfg)
    if not usuario or not clave:
        log.warning("Alerta omitida: credenciales SMTP no configuradas.")
        return False
    if not destinos:
        log.warning("Alerta omitida: no hay destinatarios de alerta configurados.")
        return False
    mensaje = MIMEText(cuerpo, "plain", "utf-8")
    mensaje["Subject"] = asunto
    mensaje["From"] = usuario
    mensaje["To"] = ", ".join(destinos)
    try:
        with smtplib.SMTP_SSL(
            correo_cfg.get("host", "smtp.gmail.com"),
            correo_cfg.get("port", 465),
        ) as servidor:
            servidor.login(usuario, clave)
            servidor.sendmail(usuario, destinos, mensaje.as_string())
        log.info(f"Alerta de fallo enviada a {len(destinos)} destinatario(s).")
        return True
    except Exception as error:  # noqa: BLE001 - la alerta nunca debe tumbar el job
        log.error(f"No se pudo enviar la alerta de fallo: {error}")
        return False


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Envia la alerta de fallo del monitor.")
    parser.add_argument("--motivo", default="fallo", choices=["fallo"])
    parser.parse_args(argv)

    cfg = cargar_config()
    umbral = 7
    try:
        umbral = int((cfg.get("umbrales", {}) or {}).get("max_dias_sin_importacion", 7))
    except (TypeError, ValueError):
        pass
    asunto, cuerpo = construir_alerta_fallo(
        dias_sin_importacion(),
        extraer_errores(),
        umbral_urgente=umbral,
    )
    enviar(asunto, cuerpo, cfg)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
