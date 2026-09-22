import json
import unittest
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import alertas


class DiasSinImportacionTests(unittest.TestCase):
    def _estado(self, directorio, contenido):
        ruta = Path(directorio) / "estado.json"
        ruta.write_text(contenido, encoding="utf-8")
        return ruta

    def test_reciente_devuelve_cero(self):
        with TemporaryDirectory() as tmp:
            ruta = self._estado(tmp, json.dumps({"ultima_revision": "2026-09-22T10:00:00+00:00"}))
            dias = alertas.dias_sin_importacion(
                ruta, ahora=datetime(2026, 9, 22, 12, 0, tzinfo=timezone.utc)
            )
            self.assertEqual(dias, 0)

    def test_antiguo_cuenta_dias(self):
        with TemporaryDirectory() as tmp:
            ruta = self._estado(tmp, json.dumps({"ultima_revision": "2026-08-25T13:44:13.915156"}))
            dias = alertas.dias_sin_importacion(
                ruta, ahora=datetime(2026, 9, 22, 12, 0, tzinfo=timezone.utc)
            )
            self.assertEqual(dias, 27)

    def test_estado_inexistente_o_corrupto_devuelve_none(self):
        with TemporaryDirectory() as tmp:
            self.assertIsNone(alertas.dias_sin_importacion(Path(tmp) / "no_existe.json"))
            ruta = self._estado(tmp, "{json invalido")
            self.assertIsNone(alertas.dias_sin_importacion(ruta))


class ConstruirAlertaTests(unittest.TestCase):
    def test_escala_a_urgente_pasado_el_umbral(self):
        asunto, cuerpo = alertas.construir_alerta_fallo(27, [], umbral_urgente=7)
        self.assertIn("[URGENTE]", asunto)
        self.assertIn("27 dia(s)", cuerpo)

    def test_aviso_normal_bajo_el_umbral(self):
        asunto, cuerpo = alertas.construir_alerta_fallo(2, [], umbral_urgente=7)
        self.assertIn("[AVISO]", asunto)
        self.assertIn("2 dia(s)", cuerpo)

    def test_sin_estado_previo_es_urgente(self):
        asunto, _ = alertas.construir_alerta_fallo(None, [], umbral_urgente=7)
        self.assertIn("[URGENTE]", asunto)

    def test_incluye_errores_y_enlace(self):
        with patch.dict(
            "os.environ",
            {
                "GITHUB_SERVER_URL": "https://github.com",
                "GITHUB_REPOSITORY": "MoneyCm/monitor-mindefensa",
                "GITHUB_RUN_ID": "123",
            },
        ):
            _, cuerpo = alertas.construir_alerta_fallo(1, ["FALLO CATASTRÓFICO: 503"], umbral_urgente=7)
        self.assertIn("503", cuerpo)
        self.assertIn("actions/runs/123", cuerpo)


class ExtraerErroresTests(unittest.TestCase):
    def test_solo_error_y_critical_en_orden(self):
        lineas = [
            json.dumps({"text": "ok", "record": {"level": {"name": "INFO"}}}),
            json.dumps({"text": "fallo 1", "record": {"level": {"name": "ERROR"}}}),
            "linea rota",
            json.dumps({"text": "fallo 2", "record": {"level": {"name": "CRITICAL"}}}),
        ]
        with TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "log.json"
            ruta.write_text("\n".join(lineas), encoding="utf-8")
            self.assertEqual(alertas.extraer_errores(ruta), ["fallo 1", "fallo 2"])

    def test_archivo_inexistente_devuelve_vacio(self):
        with TemporaryDirectory() as tmp:
            self.assertEqual(alertas.extraer_errores(Path(tmp) / "no_existe.json"), [])


class EnviarTests(unittest.TestCase):
    def test_sin_credenciales_no_envia(self):
        with patch.dict("os.environ", {}, clear=True):
            self.assertFalse(alertas.enviar("asunto", "cuerpo", {"correo": {}}))

    def test_envio_exitoso_usa_config(self):
        cfg = {"correo": {"host": "smtp.gmail.com", "port": 465, "destinatarios": ["a@x.co"]}}
        with patch.dict("os.environ", {"GMAIL_USER": "u", "GMAIL_PASS": "p"}):
            with patch("alertas.smtplib.SMTP_SSL") as smtp:
                self.assertTrue(alertas.enviar("asunto", "cuerpo", cfg))
                servidor = smtp.return_value.__enter__.return_value
                servidor.login.assert_called_once_with("u", "p")
                self.assertTrue(servidor.sendmail.called)


if __name__ == "__main__":
    unittest.main()
