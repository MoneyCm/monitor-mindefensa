import smtplib, json, os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from pathlib import Path

# Aquí el script lee los "Secrets" de GitHub en lugar de tenerlos escritos
CORREO = os.environ.get("GMAIL_USER")
PASSWD = os.environ.get("GMAIL_APP_PASSWORD")

STATE_FILE = Path("mindefensa_state.json")

def enviar_resumen():
    if not CORREO or not PASSWD:
        print("No se encontraron los Secrets de GitHub. Omitiendo correo.")
        return

    estado = {}
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as f:
            estado = json.load(f)

    ultima  = estado.get("ultima_revision", "—")
    total   = len(estado.get("archivos", {}))
    fecha_hoy = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""
    <html><body style="font-family:Calibri,Arial;background:#f4f4f8;padding:20px">
    <div style="max-width:600px;margin:0 auto;background:white;border-radius:4px;overflow:hidden">
      <div style="background:#281FD0;color:white;padding:20px 24px">
        <div style="font-size:18px;font-weight:bold;letter-spacing:1px">ALCALDÍA DE JAMUNDÍ</div>
        <div style="font-size:11px;opacity:.8;letter-spacing:2px;text-transform:uppercase">Observatorio del Delito</div>
        <div style="font-size:13px;color:#FFE000;margin-top:8px;font-weight:bold">{fecha_hoy}</div>
      </div>
      <div style="padding:24px">
        <p><b>Archivos monitoreados:</b> {total}</p>
        <p><b>Última revisión:</b> {ultima[:19]}</p>
        <p style="color:#2E7D32;font-weight:bold">✅ EJECUCIÓN EXITOSA DESDE GITHUB ACTIONS</p>
      </div>
    </div>
    </body></html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Monitor MinDefensa (GitHub) — {fecha_hoy}"
    msg["From"]    = CORREO
    msg["To"]      = CORREO
    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(CORREO, PASSWD)
        server.sendmail(CORREO, CORREO, msg.as_string())

    print(f"Correo enviado a {CORREO}")

if __name__ == "__main__":
    enviar_resumen()
