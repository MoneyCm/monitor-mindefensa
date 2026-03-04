"""
generar_reporte.py — Identidad oficial Alcaldía de Jamundí
Colores: Azul #281FD0, Amarillo #FFE000
"""
import os, sys, unicodedata
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
# Los archivos se descargan en esta carpeta según monitor.py
CARPETA  = Path("mindefensa_xlsx")
SALIDA   = "reporte_observatorio.pdf"
ESCUDO   = "escudo_jamundi.png"
MESES_ES = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

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
    "Afectación Fuerza Pública":  {"pattern": "AFECTACION FUERZA PUBLICA", "col": "CANTIDAD"},
    "Piratería Terrestre":        {"pattern": "PIRATERIA TERRESTRE",   "col": "CANTIDAD"},
    "Trata de Personas":          {"pattern": "TRATA DE PERSONAS",     "col": "CANTIDAD"},
    "Invasión de Tierras":        {"pattern": "INVASION DE TIERRAS",   "col": "CANTIDAD"},
    "Hurto a Personas":           {"pattern": "HURTO PERSONAS",        "col": "CANTIDAD"},
    "Hurto a Residencias":        {"pattern": "HURTO A RESIDENCIAS",   "col": "CANTIDAD"},
    "Hurto de Vehículos":         {"pattern": "HURTO DE VEHICULOS",    "col": "CANTIDAD"},
    "Hurto a Comercio":           {"pattern": "HURTO A COMERCIO",      "col": "CANTIDAD"},
    "Incautación Cocaína":        {"pattern": "INCAUTACION COCAINA",   "col": "CANTIDAD"},
    "Incautación Marihuana":      {"pattern": "INCAUTACION MARIHUANA", "col": "CANTIDAD"},
}

def normalizar(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto.upper()) if unicodedata.category(c) != 'Mn')

def buscar_archivo(patron):
    if not CARPETA.exists(): return None
    patron_norm = normalizar(patron)
    for f in CARPETA.glob("*.xlsx"):
        if patron_norm in normalizar(f.name):
            return f
    # También buscar en la raíz por si acaso
    for f in Path(".").glob("*.xlsx"):
        if patron_norm in normalizar(f.name):
            return f
    return None

def leer_datos():
    datos = {}
    print(f"Buscando archivos en {CARPETA.resolve()}...")
    for nombre, cfg in DATASETS_CONFIG.items():
        ruta = buscar_archivo(cfg["pattern"])
        if not ruta:
            print(f"  [AVISO] No se encontró archivo para: {nombre} (patrón: {cfg['pattern']})")
            continue
        try:
            df = pd.read_excel(ruta, engine='openpyxl')
            df.columns = [str(c).upper().strip() for c in df.columns]
            
            # Buscar columna de municipio (puede variar)
            col_muni = next((c for c in df.columns if any(x in c for x in ["COD_MUNI", "MUNICIPIO", "MPIO"])), None)
            if not col_muni:
                print(f"  [AVISO] No se encontró columna de municipio en {ruta.name}")
                continue
                
            df[col_muni] = df[col_muni].astype(str).str.strip()
            df = df[(df[col_muni] == str(COD_MUNI)) | (df[col_muni].str.contains("JAMUNDI", na=False))].copy()
            
            if df.empty:
                print(f"  [INFO] {nombre}: 0 registros para Jamundí")
                continue

            col_fecha = next((c for c in df.columns if any(x in c for x in ["FECHA_HECHO", "FECHA HECHO", "ANIO", "FECHA"])), None)
            if col_fecha:
                df['FECHA_HECHO_DT'] = pd.to_datetime(df[col_fecha], errors='coerce')
                df['ANIO'] = df['FECHA_HECHO_DT'].dt.year
                df['MES']  = df['FECHA_HECHO_DT'].dt.month
                # Fallback para años si la conversión falló pero hay columna ANIO
                if 'ANIO' in df.columns and df['ANIO'].isnull().any():
                     df['ANIO'] = pd.to_numeric(df[col_fecha].astype(str).str[:4], errors='coerce')

            # Limpiar columna de cantidad
            col_cant = cfg["col"] if cfg["col"] in df.columns else df.columns[-1]
            raw = df[col_cant]
            if raw.dtype == object:
                raw = raw.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df['col_cantidad'] = pd.to_numeric(raw, errors='coerce').fillna(0)
            
            datos[nombre] = df
            print(f"  [OK] {nombre}: {len(df)} filas encontradas para Jamundí")
        except Exception as e:
            print(f"  [ERROR] {nombre} ({ruta.name}): {e}")
    return datos

def total_anio(df, anio, hasta_mes=None):
    if 'ANIO' not in df.columns: return 0
    sub = df[df['ANIO'] == anio]
    if hasta_mes and 'MES' in df.columns:
        sub = sub[sub['MES'] <= hasta_mes]
    return float(sub['col_cantidad'].sum())

def grafica_comparativa(datos, anio_actual, anio_anterior, hasta_mes=None):
    delitos = sorted(list(datos.keys()))
    if not delitos: return None
    
    vals_ant = [total_anio(datos[d], anio_anterior, hasta_mes=hasta_mes) for d in delitos]
    vals_act = [total_anio(datos[d], anio_actual, hasta_mes=hasta_mes)   for d in delitos]
    
    fig, ax = plt.subplots(figsize=(13, 6))
    fig.patch.set_facecolor('#F4F4F8')
    ax.set_facecolor('#F4F4F8')
    
    x = range(len(delitos))
    w = 0.35
    
    ax.bar([i - w/2 for i in x], vals_ant, w, label=str(anio_anterior), color='#606175', alpha=0.8, zorder=3)
    b2 = ax.bar([i + w/2 for i in x], vals_act, w, label=str(anio_actual), color='#281FD0', alpha=0.9, zorder=3)
    
    for bar in b2:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + 0.1, f'{int(h)}', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#281FD0')

    ax.set_xticks(x)
    ax.set_xticklabels(delitos, rotation=45, ha='right', fontsize=9)
    ax.legend(frameon=False, loc='upper right')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3, zorder=0)
    
    plt.tight_layout()
    grafica_path = "comparativa_delitos.png"
    plt.savefig(grafica_path, dpi=150)
    plt.close()
    return grafica_path

def crear_pdf(datos):
    doc = SimpleDocTemplate(SALIDA, pagesize=A4, rightMargin=2*cm, leftMargin=2*cm, topMargin=2*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    
    # Estilos personalizados
    style_h1 = ParagraphStyle('H1', parent=styles['Heading1'], fontSize=18, textColor=AZUL, spaceAfter=12, fontName='Helvetica-Bold')
    style_h2 = ParagraphStyle('H2', parent=styles['Heading2'], fontSize=14, textColor=AZUL, spaceBefore=15, spaceAfter=10, fontName='Helvetica-Bold')
    style_body = ParagraphStyle('Body', parent=styles['Normal'], fontSize=10, leading=14, textColor=NEGRO, alignment=TA_LEFT)
    style_footer = ParagraphStyle('Footer', fontSize=8, textColor=GRIS, alignment=TA_CENTER)

    story = []
    hoy = datetime.now()
    anio_act = hoy.year
    anio_ant = anio_act - 1
    mes_act  = hoy.month

    # Encabezado
    try:
        if Path(ESCUDO).exists():
            img = Image(ESCUDO, width=1.8*cm, height=1.8*cm)
            img.hAlign = 'LEFT'
            story.append(img)
    except: pass
    
    story.append(Paragraph("OBSERVATORIO DEL DELITO", style_h1))
    story.append(Paragraph(f"Boletín Estadístico Mensual - Jamundí, Valle del Cauca", ParagraphStyle('Sub', fontSize=11, textColor=GRIS)))
    story.append(Paragraph(f"Fecha de generación: {hoy.strftime('%d/%m/%Y %H:%M')}", ParagraphStyle('Date', fontSize=9, textColor=GRIS, spaceAfter=20)))
    story.append(HRFlowable(width="100%", thickness=2, color=AMARILLO, spaceAfter=20))

    if not datos:
        story.append(Paragraph("No se encontraron datos para procesar el reporte en este ciclo.", style_body))
    else:
        story.append(Paragraph(f"Análisis Comparativo ({anio_ant} vs {anio_act}*)", style_h2))
        story.append(Paragraph(f"El presente reporte analiza la evolución de los principales indicadores de seguridad en el municipio de Jamundí, basándose en la información oficial consolidada por el Ministerio de Defensa Nacional.", style_body))
        story.append(Spacer(1, 10))
        
        # Gráfica
        grafica = grafica_comparativa(datos, anio_act, anio_ant, hasta_mes=mes_act)
        if grafica:
            story.append(Image(grafica, width=17*cm, height=8*cm))
            story.append(Spacer(1, 15))

        # Tabla de datos
        story.append(Paragraph("Resumen de Cifras Consolidadas", style_h2))
        
        tabla_data = [["Delito / Indicador", f"Total {anio_ant}", f"Total {anio_act}*", "Variación"]]
        for d in sorted(datos.keys()):
            v_ant = int(total_anio(datos[d], anio_ant))
            v_act = int(total_anio(datos[d], anio_act))
            var = v_act - v_ant
            var_str = f"{var:+d}" if var != 0 else "0"
            color_var = ROJO_ALT if var > 0 else VERDE if var < 0 else NEGRO
            
            tabla_data.append([
                d, 
                str(v_ant), 
                str(v_act), 
                Paragraph(f"<font color={color_var.hexval()}>{var_str}</font>", styles['Normal'])
            ])

        t = Table(tabla_data, colWidths=[8*cm, 3*cm, 3*cm, 3*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), AZUL),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('ALIGN', (0,0), (0,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 10),
            ('BOTTOMPADDING', (0,0), (-1,0), 10),
            ('BACKGROUND', (0,1), (-1,-1), GRIS_FONDO),
            ('GRID', (0,0), (-1,-1), 0.5, colors.white),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(t)
        
        story.append(Spacer(1, 25))
        story.append(Paragraph("* Datos parciales sujetos a actualización por parte de la fuente oficial.", ParagraphStyle('Note', fontSize=8, italic=True)))

    # Pie de página
    story.append(Spacer(1, 2*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=GRIS, spaceBefore=10))
    story.append(Paragraph("Fuente: Ministerio de Defensa Nacional - Información Estadística de Seguridad y Defensa.", style_footer))
    story.append(Paragraph("Generado por SISC Jamundí - Sistema de Información para la Seguridad y Convivencia.", style_footer))

    doc.build(story)
    print(f"✅ Reporte generado: {SALIDA}")

if __name__ == "__main__":
    print("Iniciando generación de reporte...")
    df_dict = leer_datos()
    crear_pdf(df_dict)
