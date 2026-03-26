"""
generar_reporte.py — Identidad oficial Alcaldía de Jamundí (Versión MinDefensa Premium)
Colores: Azul #281FD0, Amarillo #FFE000
"""

import os, sys, unicodedata, re, json
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

# ── Identidad ────────────────────────────────────────────────────────────────
AZUL       = colors.HexColor("#281FD0")
AMARILLO   = colors.HexColor("#FFE000")
NEGRO      = colors.HexColor("#1A1A2E")
GRIS       = colors.HexColor("#606175")
GRIS_FONDO = colors.HexColor("#F4F4F8")
ROJO_ALT   = colors.HexColor("#C0392B")
VERDE      = colors.HexColor("#1A7A4A")

MESES_ES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

COD_MUNI = 76364
CARPETA  = Path("mindefensa_xlsx")
SALIDA   = "reporte_observatorio.pdf"
ESCUDO   = "escudo_jamundi.png"

# Mapeo de Delito -> Fragmento de nombre de archivo y columna de datos
DATASETS_CONFIG = {
    "Homicidio Intencional":      {"pattern": "HOMICIDIO INTENCIONAL", "col": "VICTIMAS"},
    "Lesiones Comunes":           {"pattern": "LESIONES COMUNES",      "col": "CANTIDAD"},
    "Violencia Intrafamiliar":    {"pattern": "VIOLENCIA INTRAFAMILIAR", "col": "CANTIDAD"},
    "Delitos Sexuales":           {"pattern": "DELITOS SEXUALES",      "col": "CANTIDAD"},
    "Secuestro":                  {"pattern": "SECUESTRO",             "col": "CANTIDAD"},
    "Extorsión":                  {"pattern": "EXTORSION",             "col": "CANTIDAD"},
    "Terrorismo":                 {"pattern": "TERRORISMO",            "col": "CANTIDAD"},
    "Masacres":                   {"pattern": "MASACRES",              "col": "VICTIMAS"},
    "Hurto a Personas":           {"pattern": "HURTO PERSONAS",        "col": "CANTIDAD"},
    "Hurto a Residencias":        {"pattern": "HURTO A RESIDENCIAS",   "col": "CANTIDAD"},
    "Hurto de Vehículos":         {"pattern": "HURTO DE VEHICULOS",    "col": "CANTIDAD"},
    "Hurto a Comercio":           {"pattern": "HURTO A COMERCIO",      "col": "CANTIDAD"},
    "Homicidio Tránsito":         {"pattern": "HOMICIDIO ACCIDENTES DE TRANSITO", "col": "VICTIMAS"},
    "Lesiones Tránsito":          {"pattern": "LESIONES ACCIDENTES DE TRANSITO", "col": "CANTIDAD"},
    "Delitos Medio Ambiente":     {"pattern": "DELITOS CONTRA EL MEDIO AMBIENTE", "col": "CANTIDAD"},
    "Capturas Minería":           {"pattern": "CAPTURAS POR MINERIA ILEGAL", "col": "CANTIDAD"},
    "Incautaciones Minería":      {"pattern": "INCAUTACIONES MINERIA", "col": "CANTIDAD"},
    "Minería Oro/Mercurio":       {"pattern": "INCAUTACION ORO Y MERCURIO", "col": "CANTIDAD"},
    "Minas Intervenidas":         {"pattern": "MINAS INTERVENIDAS",    "col": "CANTIDAD"},
}

def normalizar(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto.upper()) if unicodedata.category(c) != 'Mn')

def buscar_archivo(patron):
    if not CARPETA.exists(): return None
    patron_norm = normalizar(patron)
    for f in CARPETA.glob("*.xlsx"):
        if patron_norm in normalizar(f.name):
            return f
    return None

def leer_datos():
    datos = {}
    print(f"📊 Buscando archivos en {CARPETA.resolve()}...")
    for nombre, cfg in DATASETS_CONFIG.items():
        ruta = buscar_archivo(cfg["pattern"])
        if not ruta: continue
        try:
            # MinDefensa Excel files often have 3-5 rows of preamble before the header
            df_full = pd.read_excel(ruta, engine='openpyxl', header=None, nrows=10)
            header_row = 0
            for i, row in df_full.iterrows():
                row_vals = [str(v).upper() for v in row.dropna()]
                if any(x in " ".join(row_vals) for x in ["COD_MUNI", "MUNICIPIO", "MPIO", "FECHA"]):
                    header_row = i
                    break
            
            df = pd.read_excel(ruta, engine='openpyxl', skiprows=header_row)
            df.columns = [str(c).upper().strip() for c in df.columns]
            col_muni = next((c for c in df.columns if any(x in c for x in ["COD_MUNI", "MUNICIPIO", "MPIO"])), None)
            if not col_muni: continue
            df[col_muni] = df[col_muni].astype(str).str.strip()
            # Filtro por código o nombre
            df = df[(df[col_muni] == str(COD_MUNI)) | (df[col_muni].str.upper().str.contains("JAMUNDI", na=False))].copy()
            if df.empty: continue

            col_fecha = next((c for c in df.columns if any(x in c for x in ["FECHA_HECHO", "FECHA HECHO", "ANIO", "FECHA"])), None)
            if col_fecha:
                df['FECHA_HECHO_DT'] = pd.to_datetime(df[col_fecha], errors='coerce')
                df['ANIO'] = df['FECHA_HECHO_DT'].dt.year
                df['MES']  = df['FECHA_HECHO_DT'].dt.month
                if 'ANIO' in df.columns and df['ANIO'].isnull().any():
                     df['ANIO'] = pd.to_numeric(df[col_fecha].astype(str).str[:4], errors='coerce')

            col_cant = cfg["col"] if cfg["col"] in df.columns else df.columns[-1]
            raw = df[col_cant]
            if raw.dtype == object:
                raw = raw.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df['col_cantidad'] = pd.to_numeric(raw, errors='coerce').fillna(0)
            datos[nombre] = df
            print(f"  [OK] {nombre}: {len(df)} registros")
        except: pass
    return datos

def total_anio(df, anio, hasta_mes=None):
    if 'ANIO' not in df.columns: return 0
    sub = df[df['ANIO'] == anio]
    if hasta_mes and 'MES' in df.columns:
        sub = sub[sub['MES'] <= hasta_mes]
    return float(sub['col_cantidad'].sum())

def fmt_val(v):
    try: return f"{int(round(float(v))):,}".replace(",", ".")
    except: return str(v)

def calcular_variacion_estado(v_prev, v_act):
    if v_prev == 0: return ("N/A", "NOVEDAD") if v_act > 0 else ("0.0%", "IGUAL")
    var = ((v_act - v_prev) / v_prev) * 100.0
    txt = f"{var:+.1f}%"
    est = "SUBE" if var > 0 else ("BAJA" if var < 0 else "IGUAL")
    return txt, est

# ── Gráficas ──
def grafica_comparativa(datos, anio_act, anio_ant, hasta_mes):
    delitos = sorted(datos.keys(), key=lambda d: total_anio(datos[d], anio_act, hasta_mes), reverse=True)[:12]
    if not delitos: return None
    v_ant = [total_anio(datos[d], anio_ant, hasta_mes) for d in delitos]
    v_act = [total_anio(datos[d], anio_act, hasta_mes) for d in delitos]
    
    fig, ax = plt.subplots(figsize=(13, 5.5), dpi=140)
    fig.patch.set_facecolor('#F4F4F8')
    ax.set_facecolor('#F4F4F8')
    
    x = range(len(delitos))
    w = 0.36
    ax.bar([i - w/2 for i in x], v_ant, w, label=str(anio_ant), color='#606175', alpha=0.85, zorder=3)
    bars_act = ax.bar([i + w/2 for i in x], v_act, w, label=str(anio_act), color='#281FD0', alpha=0.92, zorder=3)
    
    y_max = max(max(v_act) if v_act else 0, max(v_ant) if v_ant else 0, 1)
    for i, (bar, va, vb) in enumerate(zip(bars_act, v_ant, v_act)):
        if vb > 0:
            ax.text(bar.get_x() + bar.get_width()/2, vb + y_max*0.02, fmt_val(vb), 
                    ha='center', va='bottom', fontsize=8, color='#281FD0', fontweight='bold')

    ax.set_xticks(list(x))
    ax.set_xticklabels([d for d in delitos], rotation=30, ha="right", fontsize=8.5)
    ax.set_title(f"COMPARATIVO ENE-{MESES_ES[hasta_mes][:3].upper()} {anio_ant} vs {anio_act}", fontsize=11, fontweight='bold', color='#281FD0', pad=15)
    ax.legend(fontsize=9, frameon=False)
    ax.yaxis.grid(True, alpha=0.4, color='#C5C5D2', zorder=0)
    for spine in ['top','right']: ax.spines[spine].set_visible(False)
    
    plt.tight_layout()
    out = "graf_comp_mindefensa.png"
    plt.savefig(out, facecolor='#F4F4F8')
    plt.close()
    return out

def grafica_tendencia(datos, anio_act):
    plt.figure(figsize=(10, 3.5), dpi=140)
    plt.gca().set_facecolor('#F4F4F8')
    delitos_top = sorted(datos.keys(), key=lambda d: total_anio(datos[d], anio_act), reverse=True)[:5]
    if not delitos_top: return None
    for d in delitos_top:
        df = datos[d]
        if 'ANIO' in df.columns and 'MES' in df.columns:
            s = df.groupby(["ANIO", "MES"])["col_cantidad"].sum().tail(24)
            if not s.empty:
                plt.plot(range(len(s)), s.values, marker="o", label=d, linewidth=2)
    plt.title("TENDENCIA HISTÓRICA — TOP 5 INDICADORES", fontweight="bold", color="#281FD0")
    plt.legend(fontsize=7, loc="upper left", frameon=False)
    plt.grid(True, alpha=0.2)
    out = "graf_tend_mindefensa.png"
    plt.tight_layout()
    plt.savefig(out, facecolor='#F4F4F8')
    plt.close()
    return out

# ── PDF ──
def generar_pdf(datos):
    hoy = datetime.now()
    anio_act = hoy.year
    anio_ant = anio_act - 1
    # Determinar último mes con datos
    meses_con_datos = []
    for df in datos.values():
        if 'MES' in df.columns:
            mes_max = df[df['ANIO'] == anio_act]['MES'].max()
            if pd.notnull(mes_max): meses_con_datos.append(mes_max)
    mes_actual = int(max(meses_con_datos)) if meses_con_datos else hoy.month

    doc = SimpleDocTemplate(SALIDA, pagesize=A4, leftMargin=1.8*cm, rightMargin=1.8*cm, topMargin=1.2*cm, bottomMargin=1.5*cm)
    W = A4[0] - 3.6*cm
    h = []
    
    def P(txt, sz=9, b=False, c=NEGRO, a=TA_LEFT):
        return Paragraph(txt, ParagraphStyle("n", fontSize=sz, fontName="Helvetica-Bold" if b else "Helvetica", textColor=c, alignment=a))

    # Encabezado Institucional
    bloque_izq = Table([
        [P("ALCALDÍA DE JAMUNDÍ", 15, True)],
        [P("VALLE DEL CAUCA", 9, False, GRIS)],
        [Spacer(1,3)],
        [P(f"Observatorio del Delito — Boletín MinDefensa {MESES_ES[mes_actual]} {anio_act}", 9, True, AZUL)],
    ], colWidths=[W*0.62])
    
    bloque_der = Table([
        [P(datetime.now().strftime("%d/%m/%Y %H:%M"), 8, False, GRIS, TA_RIGHT)],
        [P("Secretaría de Seguridad y Convivencia", 7.5, False, GRIS, TA_RIGHT)],
        [P("Fuente: Ministerio de Defensa Nacional", 7.5, False, GRIS, TA_RIGHT)],
    ], colWidths=[W*0.38])

    enc = Table([[Image(ESCUDO, 1.4*cm, 1.9*cm) if Path(ESCUDO).exists() else "", bloque_izq, bloque_der]], colWidths=[1.7*cm, W*0.58, W*0.35])
    enc.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'), ('LEFTPADDING',(0,0),(-1,-1),0)]))
    h.append(enc)

    tl = Table([['','']], colWidths=[W*0.18, W*0.82])
    tl.setStyle(TableStyle([('BACKGROUND',(0,0),(0,0),AMARILLO), ('BACKGROUND',(1,0),(1,0),AZUL), ('ROWPADDING',(0,0),(-1,-1),3)]))
    h.append(tl)
    h.append(Spacer(1, 0.4*cm))

    # Resumen Ejecutivo
    h.append(P("RESUMEN EJECUTIVO", 11, True, AZUL))
    h.append(HRFlowable(width=W, thickness=2, color=AMARILLO))
    h.append(Spacer(1, 0.2*cm))

    filas = [[P("Indicador",8,True,colors.white,TA_CENTER), P("Corte",8,True,colors.white,TA_CENTER), P(str(anio_ant),8,True,colors.white,TA_CENTER), P(str(anio_act),8,True,colors.white,TA_CENTER), P("Variación",8,True,colors.white,TA_CENTER), P("Estado",8,True,colors.white,TA_CENTER)]]
    for d, df in sorted(datos.items(), key=lambda x: total_anio(x[1], anio_act, mes_actual), reverse=True):
        ant, act = total_anio(df, anio_ant, mes_actual), total_anio(df, anio_act, mes_actual)
        v, e = calcular_variacion_estado(ant, act)
        
        # Obtener última fecha de este delito
        corte_str = "—"
        if 'FECHA_HECHO_DT' in df.columns:
            f_max = df['FECHA_HECHO_DT'].max()
            if pd.notnull(f_max):
                corte_str = f_max.strftime("%d/%m/%y")

        clr = ROJO_ALT if e in ("SUBE", "NOVEDAD") else (VERDE if e == "BAJA" else GRIS)
        filas.append([
            P(d,8), 
            P(corte_str, 8, False, GRIS, TA_CENTER),
            P(fmt_val(ant),8,False,NEGRO,TA_CENTER), 
            P(fmt_val(act),8,True,NEGRO,TA_CENTER), 
            P(v,8,True,clr,TA_CENTER), 
            P(e,8,True,clr,TA_CENTER)
        ])

    t = Table(filas, colWidths=[W*0.28, W*0.13, W*0.11, W*0.11, W*0.19, W*0.18])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),AZUL), ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, GRIS_FONDO]), ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#C5C5D2')), ('ROWPADDING',(0,0),(-1,-1),7), ('LINEBELOW',(0,0),(-1,0),2.5,AMARILLO)]))
    h.append(t)
    h.append(Spacer(1, 0.6*cm))

    # Gráficas
    g_comp = grafica_comparativa(datos, anio_act, anio_ant, mes_actual)
    if g_comp:
        h.append(P("COMPARATIVO ANUAL POR INDICADOR", 11, True, AZUL))
        h.append(HRFlowable(width=W, thickness=2, color=AMARILLO))
        h.append(Spacer(1, 0.15*cm))
        h.append(Image(g_comp, width=W, height=W*0.38))
        h.append(Spacer(1, 0.6*cm))
        
    g_tend = grafica_tendencia(datos, anio_act)
    if g_tend:
        h.append(P("TENDENCIA HISTÓRICA — INDICADORES PRIORITARIOS", 11, True, AZUL))
        h.append(HRFlowable(width=W, thickness=2, color=AMARILLO))
        h.append(Spacer(1, 0.15*cm))
        h.append(Image(g_tend, width=W, height=W*0.35))

    # Pie
    h.append(Spacer(1, 0.6*cm))
    pie = Table([[P("ALCALDÍA DE JAMUNDÍ — SECRETARÍA DE SEGURIDAD Y CONVIVENCIA", 7.5, True, colors.white, TA_CENTER)], [P("Fuente: Ministerio de Defensa Nacional · Municipio: Jamundí (76364) · Generado automáticamente vía GitHub Actions", 7, False, colors.white, TA_CENTER)]], colWidths=[W])
    pie.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),AZUL), ('ROWPADDING',(0,0),(-1,-1),7), ('LINEABOVE',(0,0),(-1,0),3,AMARILLO)]))
    h.append(pie)

    doc.build(h)
    print(f"✅ Reporte generado: {SALIDA}")
    
    # Exportar totales y fechas de corte para el correo
    resumen = {}
    for d, df in datos.items():
        corte = "—"
        if 'FECHA_HECHO_DT' in df.columns:
            f_max = df['FECHA_HECHO_DT'].max()
            if pd.notnull(f_max):
                corte = f_max.strftime("%d/%b/%y")
        resumen[d] = {
            "valor": int(total_anio(df, anio_act, mes_actual)),
            "corte": corte
        }
    
    with open("resumen_actual.json", "w", encoding="utf-8") as f:
        json.dump(resumen, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    d = leer_datos()
    if d: generar_pdf(d)
    else: print("❌ No se encontraron datos para generar el PDF")
