import json, os
from datetime import datetime
from pathlib import Path

STATE_FILE  = Path("mindefensa_state.json")
RESUMEN_FILE = Path("resumen_actual.json")
REPORTE_PDF = Path("reporte_observatorio.pdf")

MESES_ES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio',
            'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

def generar_preview():
    estado = {}
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as f:
            estado = json.load(f)

    ultima   = estado.get("ultima_revision", "—")
    total    = len(estado.get("archivos", {}))
    nuevos   = estado.get("nuevos_ultimo", 4)  # Simulado
    cambios  = estado.get("cambios_ultimo", 2) # Simulado
    fecha_hoy = datetime.now().strftime("%d/%m/%Y %H:%M")
    mes_actual = MESES_ES[datetime.now().month]

    filas_tabla_html = ""
    if RESUMEN_FILE.exists():
        with open(RESUMEN_FILE, encoding="utf-8") as f:
            resumen_data = json.load(f)
        top_delitos = sorted(resumen_data.items(), key=lambda x: x[1]['valor'], reverse=True)[:5]
        for nombre, d in top_delitos:
            filas_tabla_html += f"""
            <tr>
              <td style="padding:8px 10px;border-bottom:1px solid #eee"><b>{nombre}</b></td>
              <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:center;color:#606175">{d['corte']}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:center;font-weight:bold;color:#281FD0">{d['valor']:,}</td>
            </tr>
            """

    color_badge = "#281FD0"
    label_badge = "VISTA PREVIA (LOCAL)"
    titulo = f"🔔 Actualización MinDefensa — {fecha_hoy}"
    descripcion = "Se han generado cambios en los datos de MinDefensa. Este es un ejemplo de cómo se verá el correo con los nuevos cortes."

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
        <h2 style="color:#281FD0;font-size:16px;margin:0 0 8px">Cordial saludo, Secretaría de Seguridad y Equipo de la Secretaría</h2>
        <p style="color:#606175;font-size:13px;margin:0 0 20px">{descripcion}</p>

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

        <p style='color:#2E7D32;font-weight:bold;font-size:13px'>✅ Se adjuntará el boletín PDF con el análisis completo.</p>
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
    with open("preview_correo.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("Previsualización generada en preview_correo.html")

if __name__ == "__main__":
    generar_preview()
