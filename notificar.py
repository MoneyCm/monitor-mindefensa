import os
import smtplib
import yaml
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from pathlib import Path
from logger import log
from utils import calculate_sha256

class Notifier:
    """Gestiona el envío de boletines vía SMTP con reporte adjunto."""
    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as f:
            self.cfg = yaml.safe_load(f)
        self.user = os.environ.get("GMAIL_USER")
        self.password = os.environ.get("GMAIL_PASS")
        self.destinatarios = self.cfg['correo']['destinatarios']
        if self.user:
            self.destinatarios.append(self.user)

    def _generar_html(self, resumen, tipo_run, pdf_sha):
        """Crea un cuerpo de correo HTML premium."""
        fecha_str = datetime.now().strftime("%d/%m/%Y %H:%M")
        rows_html = ""
        for r in sorted(resumen, key=lambda x: x['actual'], reverse=True)[:5]:
            color = "#C0392B" if r['estado'] == "SUBE" else ("#2E7D32" if r['estado'] == "BAJA" else "#606175")
            ult_reg = r.get('ultimo_registro', 'N/A')
            barrios = r.get('barrios_top', 'N/A')
            rows_html += f"""
            <tr>
                <td style="padding:10px; border-bottom:1px solid #eee;"><b>{r['delito']}</b></td>
                <td style="padding:10px; border-bottom:1px solid #eee; text-align:center;">{r['actual']}</td>
                <td style="padding:10px; border-bottom:1px solid #eee; text-align:center; color:{color}; font-weight:bold;">{r['variacion']}</td>
                <td style="padding:10px; border-bottom:1px solid #eee; text-align:center; font-size: 11px;">{ult_reg}</td>
                <td style="padding:10px; border-bottom:1px solid #eee; font-size: 11px; color:#555;">{barrios}</td>
            </tr>
            """

        asunto_prefijo = self.cfg['correo']['prefijo_reunion'] if tipo_run == "reunion" else \
                         (self.cfg['correo']['prefijo_consejo'] if tipo_run == "consejo" else "🔔 Actualización")
        
        titulo = f"Boletín MinDefensa - Observatorio Jamundí"

        html = f"""
        <html>
        <body style="font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f4f8; margin: 0; padding: 20px;">
            <div style="max-width: 650px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
                <div style="background: {self.cfg['estetica']['azul']}; padding: 30px; color: white;">
                    <div style="font-size: 10px; letter-spacing: 2px; text-transform: uppercase; color: {self.cfg['estetica']['amarillo']}; font-weight: bold;">Alcaldía de Jamundí</div>
                    <h1 style="margin: 10px 0 0; font-size: 22px;">{titulo}</h1>
                    <div style="font-size: 12px; margin-top: 5px; opacity: 0.8;">{fecha_str} | Generado Automáticamente</div>
                </div>
                <div style="padding: 30px;">
                    <p style="color: #444; font-size: 15px; line-height: 1.6;">Cordial saludo,<br><br>Se adjunta el <b>Boletín de Seguridad Institucional (Fuente: MinDefensa)</b> con el análisis consolidado de las últimas actualizaciones de bases de datos nacionales.</p>
                    
                    <h3 style="color: {self.cfg['estetica']['azul']}; font-size: 14px; text-transform: uppercase;">Resumen de Indicadores Top</h3>
                    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                        <tr style="background: #f8f9fa;">
                            <th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd; width: 25%;">Indicador</th>
                            <th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd; width: 12%;">Actual</th>
                            <th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd; width: 12%;">Var.</th>
                            <th style="padding: 10px; text-align: center; border-bottom: 2px solid #ddd; width: 16%;">Últ. Reg.</th>
                            <th style="padding: 10px; text-align: left; border-bottom: 2px solid #ddd; width: 35%;">Zonas Calientes (2026)</th>
                        </tr>
                        {rows_html}
                    </table>

                    <div style="margin-top: 30px; padding: 15px; background: #fffde7; border-left: 4px solid {self.cfg['estetica']['amarillo']}; font-size: 12px; color: #555;">
                        <b>Integridad del Reporte:</b><br>
                        Checksum SHA256: <code style="font-size: 10px;">{pdf_sha}</code>
                    </div>

                    <p style="color: #666; font-size: 13px; margin-top: 30px;">Para un análisis detallado, por favor consulte el archivo PDF adjunto.</p>
                    
                    <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 15px;">
                        <p style="margin: 0; font-size: 14px; font-weight: bold; color: {self.cfg['estetica']['azul']};">Elaborado por:</p>
                        <p style="margin: 5px 0 0; font-size: 13px; color: #555;">César Alfonso Forero Molano<br>Profesional Secretaría de Seguridad y Convivencia</p>
                    </div>
                </div>
                <div style="background: #f8f9fa; padding: 20px; text-align: center; font-size: 11px; color: #999; border-top: 1px solid #eee;">
                    Secretaría de Seguridad y Convivencia | Jamundí de Cara a la Gente
                </div>
            </div>
        </body>
        </html>
        """
        return asunto_prefijo, html

    def enviar(self, pdf_path, tipo_run="normal"):
        """Envía el correo electrónico si las credenciales están presentes."""
        if not self.user or not self.password:
            log.warning("Credenciales SMTP no configuradas. Saltando envío de correo.")
            return False

        # Cargar resumen
        resumen = []
        fecha_registro = ""
        resumen_path = Path("resumen_actual.json")
        if resumen_path.exists():
            with open(resumen_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    resumen = data.get("indicadores", [])
                    fecha_registro = f"{data.get('mes_corte', '')} {data.get('anio_act', '')}"
                else:
                    resumen = data

        sha = calculate_sha256(pdf_path)
        prefijo, html = self._generar_html(resumen, tipo_run, sha)
        
        msg = MIMEMultipart()
        asunto = "Boletín MinDefensa"
        if fecha_registro:
            asunto += f" - {fecha_registro.strip()}"
        msg['Subject'] = asunto
        msg['From'] = self.user
        msg['To'] = ", ".join(self.destinatarios)

        msg.attach(MIMEText(html, 'html'))

        # Adjuntar PDF
        if Path(pdf_path).exists():
            with open(pdf_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            filename = f"Bol_Seguridad_Jamundi_{datetime.now().strftime('%Y%m%d')}.pdf"
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        try:
            with smtplib.SMTP_SSL(self.cfg['correo']['host'], self.cfg['correo']['port']) as server:
                server.login(self.user, self.password)
                server.sendmail(self.user, self.destinatarios, msg.as_string())
            log.info(f"Correo enviado exitosamente a {len(self.destinatarios)} destinatarios.")
            return True
        except Exception as e:
            log.error(f"Error enviando correo: {e}")
            return False
