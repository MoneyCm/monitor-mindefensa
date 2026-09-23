import unittest

from notificar import Notifier


def _notificador():
    notificador = Notifier.__new__(Notifier)
    notificador.cfg = {
        "estetica": {"azul": "#281FD0", "amarillo": "#FFE000"},
        "correo": {},
    }
    return notificador


class GenerarHtmlTests(unittest.TestCase):
    def test_acepta_resumen_con_varpct(self):
        resumen = [
            {
                "delito": "Homicidios",
                "actual": 81,
                "varPct": "+3.0%",
                "estado": "SUBE",
                "ultimo_registro": "30/08/2026",
                "barrios_top": "Zona X",
            }
        ]
        _, html = _notificador()._generar_html(resumen, "normal", "abc123")
        self.assertIn("Homicidios", html)
        self.assertIn("+3.0%", html)

    def test_acepta_resumen_con_variacion_legada(self):
        resumen = [
            {
                "delito": "Secuestro",
                "actual": 4,
                "variacion": "+100.0%",
                "estado": "SUBE",
                "ultimo_registro": "24/02/2026",
                "barrios_top": "Sin datos de barrio",
            }
        ]
        _, html = _notificador()._generar_html(resumen, "normal", "abc123")
        self.assertIn("+100.0%", html)

    def test_sin_clave_de_variacion_no_falla(self):
        resumen = [
            {
                "delito": "Masacres",
                "actual": 0,
                "estado": "BAJA",
                "ultimo_registro": "10/06/2025",
                "barrios_top": "Sin datos de barrio",
            }
        ]
        _, html = _notificador()._generar_html(resumen, "normal", "abc123")
        self.assertIn("Masacres", html)


if __name__ == "__main__":
    unittest.main()
