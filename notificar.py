import smtplib, json, os, sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from pathlib import Path

CORREO = os.environ.get("GMAIL_USER")
PASSWD = os.environ.get("GMAIL_APP_PASSWORD")
STATE_FILE  = Path("mindefensa_state.json")
REPORTE_PDF = Path("reporte_observatorio.pdf")

MESES_ES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio',
            'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

def tipo_envio():
    """Determina el tipo de envío según el día de la semana."""
    dia = datetime.now().weekday()  # 0=lunes, 1=martes, ... 4=viernes
    if dia == 1:   # martes → previo reunión miércoles
        return "reunion"
    elif dia == 4: # viernes → previo consejo seguridad lunes
        return "consejo"
    else:
        return "cambio"

def asunto_y_titulo(tipo, fecha_hoy):
    if tipo == "reunion":
        return (
            f"📋 Boletín Seguridad — Reunión de Planeación {fecha_hoy}",
            "Boletín previo a la Reunión de Planeación de Seguridad",
            "Este boletín ha sido generado automáticamente para la reunión del miércoles."
        )
    elif tipo == "consejo":
        return (
            f"🔒 Boletín Seguridad — Consejo de Seguridad {fecha_hoy}",
            "Boletín previo al Consejo Municipal de Seguridad",
            "Este boletín ha sido generado automáticamente para el Consejo de Seguridad del lunes."
        )
    else:
        return (
            f"🔔 Actualización MinDefensa — {fecha_hoy}",
            "Actualización detectada en datos de MinDefensa",
            "Se detectaron cambios en los archivos de MinDefensa. Se adjunta el boletín actualizado."
        )

def enviar_resumen():
    if not CORREO or not PASSWD:
        print("No se encontraron los Secrets de GitHub. Omitiendo correo.")
        return

    estado = {}
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as f:
            estado = json.load(f)

    ultima   = estado.get("ultima_revision", "—")
    total    = len(estado.get("archivos", {}))
    nuevos   = estado.get("nuevos_ultimo", 0)
    cambios  = estado.get("cambios_ultimo", 0)
    fecha_hoy = datetime.now().strftime("%d/%m/%Y %H:%M")
    mes_actual = MESES_ES[datetime.now().month]

    # Cargar detalle con fechas de corte
    filas_tabla_html = ""
    RESUMEN_FILE = Path("resumen_actual.json")
    if RESUMEN_FILE.exists():
        try:
            with open(RESUMEN_FILE, encoding="utf-8") as f:
                resumen_data = json.load(f)
            # Mostrar top 5 delitos por volumen
            top_delitos = sorted(resumen_data.items(), key=lambda x: x[1]['valor'], reverse=True)[:5]
            for nombre, d in top_delitos:
                filas_tabla_html += f"""
                <tr>
                  <td style="padding:8px 10px;border-bottom:1px solid #eee"><b>{nombre}</b></td>
                  <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:center;color:#606175">{d['corte']}</td>
                  <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:center;font-weight:bold;color:#281FD0">{d['valor']:,}</td>
                </tr>
                """
        except: 
            filas_tabla_html = "<tr><td colspan='3' style='padding:10px;text-align:center'>Error al cargar detalle</td></tr>"
    else:
        filas_tabla_html = "<tr><td colspan='3' style='padding:10px;text-align:center'>No hay detalle disponible</td></tr>"

    tipo = tipo_envio()
    asunto, titulo, descripcion = asunto_y_titulo(tipo, fecha_hoy)

    # Badge color según tipo
    color_badge = {"reunion": "#FFB600", "consejo": "#C0392B", "cambio": "#281FD0"}[tipo]
    label_badge = {"reunion": "REUNIÓN MIÉRCOLES", "consejo": "CONSEJO SEGURIDAD", "cambio": "ACTUALIZACIÓN"}[tipo]

    html = f"""
    <html><body style="font-family:Calibri,Arial;background:#f4f4f8;padding:20px;margin:0">
    <div style="max-width:620px;margin:0 auto;background:white;border-radius:6px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1)">

      <!-- Header azul -->
      <div style="background:#281FD0;padding:24px 28px 16px">
        <div style="font-size:11px;color:#FFE000;letter-spacing:3px;font-weight:bold;text-transform:uppercase">Alcaldía de Jamundí · Valle del Cauca</div>
        <div style="font-size:20px;font-weight:bold;color:white;margin-top:4px">Observatorio del Delito</div>
        <div style="font-size:12px;color:rgba(255,255,255,.7);margin-top:4px">{fecha_hoy} · GitHub Actions</div>
      </div>

      <!-- Línea amarilla -->
      <div style="height:4px;background:linear-gradient(to right,#FFE000 20%,#281FD0 20%)"></div>

      <!-- Badge tipo -->
      <div style="padding:16px 28px 0">
        <span style="background:{color_badge};color:white;font-size:10px;font-weight:bold;letter-spacing:2px;padding:4px 12px;border-radius:20px">{label_badge}</span>
      </div>

      <!-- Contenido -->
      <div style="padding:20px 28px">
        <h2 style="color:#281FD0;font-size:16px;margin:0 0 8px">Cordial saludo, Sr. Secretario de Seguridad y Equipo de la Secretaría</h2>
        <p style="color:#606175;font-size:13px;margin:0 0 20px">Por medio de la presente, presento el <b>Boletín de Seguimiento de Indicadores de Seguridad</b> generado automáticamente a partir de las últimas actualizaciones de la plataforma oficial de MinDefensa (con fecha de corte 24 de marzo de 2026).</p>
        <p style="color:#606175;font-size:13px;margin:0 0 20px">A continuación, presentamos los indicadores prioritarios para el municipio de Jamundí en el periodo analizado:</p>

        <!-- Stats Globales -->
        <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
          <tr>
            <td style="background:#f4f4f8;padding:12px;border-radius:4px;text-align:center;width:33%">
              <div style="font-size:24px;font-weight:bold;color:#281FD0">{total}</div>
              <div style="font-size:11px;color:#606175">Archivos monitoreados</div>
            </td>
            <td style="width:2%"></td>
            <td style="background:#f4f4f8;padding:12px;border-radius:4px;text-align:center;width:33%">
              <div style="font-size:24px;font-weight:bold;color:#{"C0392B" if nuevos > 0 else "2E7D32"}">{nuevos}</div>
              <div style="font-size:11px;color:#606175">Archivos nuevos</div>
            </td>
            <td style="width:2%"></td>
            <td style="background:#f4f4f8;padding:12px;border-radius:4px;text-align:center;width:33%">
              <div style="font-size:24px;font-weight:bold;color:#{"FFB600" if cambios > 0 else "2E7D32"}">{cambios}</div>
              <div style="font-size:11px;color:#606175">Archivos actualizados</div>
            </td>
          </tr>
        </table>

        <!-- Detalle de Delitos con Corte -->
        <h3 style="color:#281FD0;font-size:13px;margin:20px 0 10px;text-transform:uppercase;letter-spacing:1px">Resumen de Indicadores (Jamundí)</h3>
        <table style="width:100%;border-collapse:collapse;font-size:12px;background:#f8f9fa;border-radius:8px">
          <tr style="background:#e9ecef">
            <th style="padding:10px;text-align:left;border-bottom:2px solid #dee2e6">Delito</th>
            <th style="padding:10px;text-align:center;border-bottom:2px solid #dee2e6">Corte</th>
            <th style="padding:10px;text-align:center;border-bottom:2px solid #dee2e6">Casos</th>
          </tr>
          {filas_tabla_html}
        </table>
        <p style="font-size:11px;color:#858796;margin-top:10px"><i>* Las fechas reflejan el último registro reportado oficialmente.</i></p>

        <p style="color:#606175;font-size:12px">
          <b>Última revisión:</b> {ultima[:19] if ultima != "—" else "—"}<br>
          <b>Período del boletín:</b> {mes_actual} {datetime.now().year}
        </p>

        {"<p style='color:#2E7D32;font-weight:bold;font-size:13px'>✅ Se adjunta el boletín PDF con el análisis completo.</p>" if REPORTE_PDF.exists() else ""}
        <br>
        <p style="color:#606175;font-size:13px;margin:0">Atentamente,</p>
        <p style="color:#281FD0;font-size:14px;margin:0;font-weight:bold">César Forero</p>
        <p style="color:#606175;font-size:12px;margin:0">Observatorio del Delito</p>
        <p style="color:#606175;font-size:12px;margin:0">Secretaría de Seguridad y Convivencia</p>
      </div>

      <!-- Footer -->
      <div style="background:#281FD0;padding:12px 28px;border-top:3px solid #FFE000">
        <p style="color:rgba(255,255,255,.7);font-size:10px;margin:0;text-align:center">
          Alcaldía de Jamundí · Secretaría de Seguridad y Convivencia · www.jamundi.gov.co<br>
          Fuente: Ministerio de Defensa Nacional · Código municipio: 76364
        </p>
      </div>
    </div>
    </body></html>
    """

    msg = MIMEMultipart("mixed")
    msg["Subject"] = asunto
    msg["From"]    = CORREO
    msg["To"]      = f"{CORREO}, secretaria.seguridad@jamundi.gov.co"
    msg.attach(MIMEText(html, "html"))

    # Adjuntar PDF si existe
    if REPORTE_PDF.exists():
        with open(REPORTE_PDF, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        nombre_pdf = f"Boletin_Observatorio_Jamundi_{datetime.now().strftime('%Y%m%d')}.pdf"
        part.add_header("Content-Disposition", f"attachment; filename={nombre_pdf}")
        msg.attach(part)
        print(f"PDF adjunto: {nombre_pdf}")
    else:
        print("PDF no encontrado, enviando sin adjunto")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(CORREO, PASSWD)
        destinatarios = [CORREO, "secretaria.seguridad@jamundi.gov.co"]
        server.sendmail(CORREO, destinatarios, msg.as_string())
    print(f"Correo enviado a {destinatarios} [{tipo}]")

if __name__ == "__main__":
    enviar_resumen()
