"""
generar_reporte.py — Identidad oficial Alcaldía de Jamundí
Colores: Azul #281FD0, Amarillo #FFE000
"""
import os, sys
from datetime import datetime
from pathlib import Path
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable

AZUL       = colors.HexColor("#281FD0")
AZUL_CLARO = colors.HexColor("#3A30F1")
AMARILLO   = colors.HexColor("#FFE000")
NEGRO      = colors.HexColor("#1A1A2E")
GRIS       = colors.HexColor("#606175")
GRIS_FONDO = colors.HexColor("#F4F4F8")
ROJO_ALT   = colors.HexColor("#C0392B")
VERDE      = colors.HexColor("#1A7A4A")

COD_MUNI = 76364
CARPETA  = "."
SALIDA   = "reporte_observatorio.pdf"
ESCUDO   = "escudo_jamundi.png"
MESES_ES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

DATASETS = {
    "Homicidio Intencional":      {"file": "HOMICIDIO INTENCIONAL.xlsx",                               "col": "VICTIMAS"},
    "Lesiones Comunes":           {"file": "LESIONES COMUNES.xlsx",                                    "col": "CANTIDAD"},
    "Violencia Intrafamiliar":    {"file": "VIOLENCIA INTRAFAMILIAR.xlsx",                             "col": "CANTIDAD"},
    "Delitos Sexuales":           {"file": "DELITOS SEXUALES.xlsx",                                    "col": "CANTIDAD"},
    "Secuestro":                  {"file": "SECUESTRO.xlsx",                                           "col": "CANTIDAD"},
    "Extorsion":                  {"file": "EXTORSIÓN.xlsx",                                      "col": "CANTIDAD"},
    "Terrorismo":                 {"file": "TERRORISMO.xlsx",                                          "col": "CANTIDAD"},
    "Masacres":                   {"file": "MASACRES.xlsx",                                            "col": "VICTIMAS"},
    "Afectacion Fuerza Publica":  {"file": "AFECTACIÓN A LA FUERZA PÚBLICA.xlsx",            "col": "CANTIDAD"},
    "Pirateria Terrestre":        {"file": "PIRATERÍA TERRESTRE.xlsx",                            "col": "CANTIDAD"},
    "Trata de Personas":          {"file": "TRATA DE PERSONAS Y TRÁFICO DE MIGRANTES.xlsx",       "col": "CANTIDAD"},
    "Invasion de Tierras":        {"file": "INVASIÓN DE TIERRAS.xlsx",                            "col": "CANTIDAD"},
    "Hurto a Personas":           {"file": "HURTO PERSONAS.xlsx",                                      "col": "CANTIDAD"},
    "Hurto a Residencias":        {"file": "HURTO A RESIDENCIAS.xlsx",                                 "col": "CANTIDAD"},
    "Hurto de Vehiculos":         {"file": "HURTO DE VEHÍCULOS.xlsx",                             "col": "CANTIDAD"},
    "Hurto a Comercio":           {"file": "HURTO A COMERCIO.xlsx",                                    "col": "CANTIDAD"},
    "Incautacion Cocaina":        {"file": "INCAUTACIÓN DE COCAINA.xlsx",                         "col": "CANTIDAD"},
    "Incautacion Marihuana":      {"file": "INCAUTACIÓN DE MARIHUANA.xlsx",                       "col": "CANTIDAD"},
}

def leer_datos():
    datos = {}
    for nombre, cfg in DATASETS.items():
        ruta = Path(CARPETA) / cfg["file"]
        if not ruta.exists():
            print(f"  [AVISO] No se encontro {ruta}")
            continue
        try:
            df = pd.read_excel(ruta, engine='openpyxl')
            df.columns = [str(c).upper().strip() for c in df.columns]
            df['COD_MUNI'] = pd.to_numeric(df['COD_MUNI'], errors='coerce')
            df = df[df['COD_MUNI'] == COD_MUNI].copy()
            col_fecha = 'FECHA_HECHO' if 'FECHA_HECHO' in df.columns else 'FECHA HECHO'
            df['FECHA_HECHO'] = pd.to_datetime(df[col_fecha], errors='coerce')
            df['ANIO'] = df['FECHA_HECHO'].dt.year
            df['MES']  = df['FECHA_HECHO'].dt.month
            raw = df[cfg['col']]
            if raw.dtype == object:
                raw = raw.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df['col_cantidad'] = pd.to_numeric(raw, errors='coerce').fillna(0)
            datos[nombre] = df
            print(f"  [OK] {nombre}: {len(df)} filas")
        except Exception as e:
            print(f"  [ERROR] {nombre}: {e}")
    return datos

def total_anio(df, anio, hasta_mes=None):
    sub = df[df['ANIO'] == anio]
    if hasta_mes:
        sub = sub[sub['MES'] <= hasta_mes]
    return float(sub['col_cantidad'].sum())

def grafica_comparativa(datos, anio_actual, anio_anterior, hasta_mes=None):
    delitos  = list(datos.keys())
    vals_ant = [total_anio(datos[d], anio_anterior, hasta_mes=hasta_mes) for d in delitos]
    vals_act = [total_anio(datos[d], anio_actual, hasta_mes=hasta_mes)   for d in delitos]
    fig, ax = plt.subplots(figsize=(13, 5))
    fig.patch.set_facecolor('#F4F4F8')
    ax.set_facecolor('#F4F4F8')
    x = range(len(delitos))
    w = 0.36
    ax.bar([i - w/2 for i in x], vals_ant, w, label=str(anio_anterior), color='#606175', alpha=0.85, zorder=3)
    b2 = ax.bar([i + w/2 for i in x], vals_act, w, label=str(anio_actual), color='#281FD0', alpha=0.92, zorder=3)
    for bar in b2:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + 0.3, str(int(h)), ha='center', va='bottom', fontsize=8, color='#281FD0', fontweight='bold')
    ax.set_xticks(list(x))
    ax.set_xticklabels([d.replace(' ', '\n') for d in delitos], fontsize=8.5, color='#1A1A2E')
    ax.set_ylabel('Casos / kg', fontsize=9, color='#606175')
    ax.set_title(f'Comparativo ene-{MESES_ES[hasta_mes][:3] if hasta_mes else "Dic"} {anio_anterior} vs {anio_actual}', fontsize=11, fontweight='bold', color='#281FD0', pad=10)
    ax.legend(fontsize=9, framealpha=0.7)
    for spine in ['top','right']:
        ax.spines[spine].set_visible(False)
    ax.yaxis.grid(True, alpha=0.4, color='#C5C5D2', zorder=0)
    ax.set_axisbelow(True)
    ax.axhline(y=0, color='#FFE000', linewidth=2.5, zorder=4)
    plt.tight_layout()
    ruta = 'graf_comp.png'
    plt.savefig(ruta, dpi=150, bbox_inches='tight', facecolor='#F4F4F8')
    plt.close()
    return ruta

def grafica_tendencia_smallmultiples(datos, fecha_max, delitos, meses_back=24):
    fin = pd.Timestamp(fecha_max).to_period("M").to_timestamp("M")
    ini = (fin.to_period("M") - (meses_back - 1)).to_timestamp("M")
    meses = pd.period_range(ini.to_period("M"), fin.to_period("M"), freq="M").to_timestamp("M")
    n = len(delitos)
    fig = plt.figure(figsize=(10, 5.2))
    gs = fig.add_gridspec(nrows=n, ncols=1, hspace=0.15)
    max_y = 0
    series = {}
    for d in delitos:
        df = datos.get(d)
        if df is None or len(df) == 0:
            y = [0]*len(meses)
        else:
            sub = df.copy()
            sub = sub[(sub["FECHA_HECHO"] >= ini) & (sub["FECHA_HECHO"] <= fin)]
            g = (sub.groupby(sub["FECHA_HECHO"].dt.to_period("M"))["col_cantidad"].sum().reindex(meses.to_period("M"), fill_value=0))
            y = g.values.tolist()
        y2 = [float(v) for v in y]
        series[d] = y2
        max_y = max(max_y, max(y2) if y2 else 0)
    for i, d in enumerate(delitos):
        ax = fig.add_subplot(gs[i, 0])
        y = series[d]
        ax.plot(meses, y, linewidth=1.8, alpha=0.95)
        try:
            ax.text(meses[-1], y[-1], f" {int(round(y[-1]))}", fontsize=8, va="center")
        except: pass
        ax.grid(True, linewidth=0.3, alpha=0.25)
        ax.text(0.01, 0.85, d, transform=ax.transAxes, fontsize=9, fontweight="bold")
        ax.set_ylim(0, max(1, max_y * 1.05))
        if i < n - 1: ax.set_xticklabels([])
        else:
            ticks = meses[::3]
            ax.set_xticks(ticks)
            ax.set_xticklabels([m.strftime("%b-%y") for m in ticks], fontsize=8, rotation=35, ha="right")
    fig.suptitle(f"Tendencia mensual (últimos {meses_back} meses) — Jamundí", fontsize=11, y=0.98)
    out = "tendencia_24m.png"
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    fig.savefig(out, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return str(out)

def fmt_val(nombre, v):
    if v is None: v = 0
    if "Incautacion" in nombre or "Incautación" in nombre:
        try:
            num = float(v)
            txt = f"{num:.3f}".rstrip("0").rstrip(".")
            return f"{txt} kg"
        except: return f"{v} kg"
    try: return f"{int(round(float(v)))}"
    except: return str(v)

def calcular_variacion_estado(v_prev, v_act):
    if v_prev == 0 and v_act == 0: return "0.0%", "IGUAL"
    if v_prev == 0 and v_act > 0: return "N/A", "APARECE"
    if v_prev > 0 and v_act == 0: return "-100.0%", "BAJA"
    var = ((v_act - v_prev) / v_prev) * 100.0
    return (f"+{var:.1f}%" if var > 0 else f"{var:.1f}%"), ("SUBE" if var > 0 else "BAJA" if var < 0 else "IGUAL")

def generar_pdf(datos, ruta_salida):
    hoy = datetime.now()
    anio_actual = hoy.year
    anio_anterior = anio_actual - 1
    fecha_max = max(df['FECHA_HECHO'].max() for df in datos.values())
    mes_actual = (fecha_max.month - 1 if fecha_max.day < 25 and fecha_max.month > 1 else 12 if fecha_max.day < 25 else fecha_max.month)
    mes_nombre = MESES_ES[mes_actual]
    doc = SimpleDocTemplate(ruta_salida, pagesize=A4, leftMargin=1.8*cm, rightMargin=1.8*cm, topMargin=1.2*cm, bottomMargin=1.5*cm)
    W = A4[0] - 3.6*cm
    estilos = getSampleStyleSheet()
    def E(n, **kw): return ParagraphStyle(n, parent=estilos['Normal'], **kw)
    historia = []
    bloque_nombre = Table([[Paragraph("ALCALDÍA DE JAMUNDÍ", E('i', fontSize=15, fontName='Helvetica-Bold', textColor=NEGRO))], [Paragraph("VALLE DEL CAUCA", E('d', fontSize=9, textColor=GRIS))], [Spacer(1,3)], [Paragraph(f"Observatorio del Delito — Boletín {mes_nombre} {anio_actual}", E('o', fontSize=9, fontName='Helvetica-Bold', textColor=AZUL_CLARO))]], colWidths=[W*0.62])
    bloque_fecha = Table([[Paragraph(hoy.strftime('%d/%m/%Y %H:%M'), E('f1', fontSize=8, textColor=GRIS, alignment=TA_RIGHT))], [Paragraph("Secretaría de Seguridad y Convivencia", E('f2', fontSize=7.5, textColor=GRIS, alignment=TA_RIGHT))], [Paragraph("Fuente: Ministerio de Defensa Nacional", E('f3', fontSize=7.5, textColor=GRIS, alignment=TA_RIGHT))]], colWidths=[W*0.38])
    enc_data = [[Image(ESCUDO, width=1.4*cm, height=1.9*cm) if os.path.exists(ESCUDO) else "", bloque_nombre, bloque_fecha]]
    historia.append(Table(enc_data, colWidths=[1.7*cm, W*0.60, W*0.40 - 1.7*cm] if os.path.exists(ESCUDO) else [W*0.62, W*0.38]))
    historia.append(Spacer(1, 0.4*cm))
    historia.append(Paragraph("RESUMEN EJECUTIVO", E('h2', fontSize=11, fontName='Helvetica-Bold', textColor=AZUL)))
    historia.append(HRFlowable(width=W, thickness=2, color=AMARILLO))
    historia.append(Spacer(1, 0.2*cm))
    def TH(t): return Paragraph(f"<b>{t}</b>", E('th', fontSize=8, fontName='Helvetica-Bold', textColor=colors.white, alignment=TA_CENTER))
    def TD(t, bold=False, clr=NEGRO, align=TA_LEFT): return Paragraph(f"<b>{t}</b>" if bold else str(t), E('td', fontSize=8.5, textColor=clr, alignment=align))
    filas = [[TH("Delito"), TH(str(anio_anterior)), TH(str(anio_actual)), TH("Variación"), TH("Estado")]]
    for nombre, df in datos.items():
        ant, act = total_anio(df, anio_anterior, hasta_mes=mes_actual), total_anio(df, anio_actual, hasta_mes=mes_actual)
        var_txt, estado_txt = calcular_variacion_estado(ant, act)
        clr_v = ROJO_ALT if (act-ant) > 0 else (VERDE if (act-ant) < 0 else GRIS)
        filas.append([TD(nombre), TD(fmt_val(nombre, ant), align=TA_CENTER), TD(fmt_val(nombre, act), bold=True, align=TA_CENTER), TD(var_txt, clr=clr_v, align=TA_CENTER, bold=True), TD(estado_txt, clr=clr_v, align=TA_CENTER, bold=True)])
    tabla_res = Table(filas, colWidths=[W*0.38, W*0.13, W*0.13, W*0.20, W*0.16], repeatRows=1)
    tabla_res.setStyle(TableStyle([('BACKGROUND', (0,0),(-1,0), AZUL), ('ROWBACKGROUNDS',(0,1),(-1,-1), [colors.white, GRIS_FONDO]), ('GRID', (0,0),(-1,-1), 0.4, colors.HexColor('#C5C5D2')), ('ROWPADDING', (0,0),(-1,-1), 7), ('VALIGN', (0,0),(-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0),(-1,0), 2.5, AMARILLO)]))
    historia.append(tabla_res)
    historia.append(Spacer(1, 0.5*cm))
    historia.append(Paragraph("COMPARATIVO ANUAL POR DELITO", E('h3', fontSize=11, fontName='Helvetica-Bold', textColor=AZUL)))
    historia.append(HRFlowable(width=W, thickness=2, color=AMARILLO))
    try: historia.append(Image(grafica_comparativa(datos, anio_actual, anio_anterior, hasta_mes=mes_actual), width=W, height=W*0.38))
    except Exception as e: historia.append(Paragraph(f"Grafica no disponible: {e}", E('err', fontSize=8)))
    historia.append(Spacer(1, 0.4*cm))
    historia.append(Paragraph("TENDENCIA MENSUAL — DELITOS PRIORITARIOS", E('h4', fontSize=11, fontName='Helvetica-Bold', textColor=AZUL)))
    historia.append(HRFlowable(width=W, thickness=2, color=AMARILLO))
    try:
        delitos_prio = [d for d in ['Homicidio Intencional','Lesiones Comunes','Violencia Intrafamiliar'] if d in datos]
        if delitos_prio: historia.append(Image(grafica_tendencia_smallmultiples(datos, fecha_max, delitos_prio, meses_back=24), width=W, height=W*0.55))
    except Exception as e: historia.append(Paragraph(f"Grafica no disponible: {e}", E('err2', fontSize=8)))
    doc.build(historia)
    print(f"\n REPORTE GENERADO: {ruta_salida}")

if __name__ == '__main__':
    print("=" * 60); print("OBSERVATORIO DEL DELITO — ALCALDIA DE JAMUNDI"); print("=" * 60)
    datos = leer_datos()
    if not datos: print("No se encontraron datos."); sys.exit(1)
    generar_pdf(datos, sys.argv[1] if len(sys.argv) > 1 else SALIDA)
