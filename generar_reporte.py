import yaml
from pathlib import Path
from datetime import datetime
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
from logger import log

class PDFGenerator:
    """Generador de boletín PDF institucional de alta calidad."""
    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as f:
            self.cfg = yaml.safe_load(f)
        self.azul = colors.HexColor(self.cfg['estetica']['azul'])
        self.amarillo = colors.HexColor(self.cfg['estetica']['amarillo'])
        self.gris_fondo = colors.HexColor(self.cfg['estetica']['graficas_fondo'])
        self.meses_es = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

    def _crear_grafica_comparativa(self, resumen_datos, anio_act, anio_ant, mes_hasta):
        """Genera una gráfica comparativa de delitos top."""
        # Filtrar solo delitos con datos
        delitos = sorted(resumen_datos.keys(), key=lambda d: resumen_datos[d]['actual'], reverse=True)[:10]
        if not delitos: return None

        v_ant = [resumen_datos[d]['anterior'] for d in delitos]
        v_act = [resumen_datos[d]['actual'] for d in delitos]
        
        fig, ax = plt.subplots(figsize=(10, 5), dpi=120)
        fig.patch.set_facecolor(self.cfg['estetica']['graficas_fondo'])
        ax.set_facecolor(self.cfg['estetica']['graficas_fondo'])
        
        x = range(len(delitos))
        width = 0.35
        
        ax.bar([i - width/2 for i in x], v_ant, width, label=str(anio_ant), color='#606175', alpha=0.8)
        bars = ax.bar([i + width/2 for i in x], v_act, width, label=str(anio_act), color='#281FD0')
        
        # Etiquetas de valores
        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'{int(height)}', xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=8, fontweight='bold')

        ax.set_xticks(x)
        ax.set_xticklabels(delitos, rotation=25, ha="right", fontsize=9)
        ax.legend()
        ax.set_title(f"Comparativo Acumulado Ene-{self.meses_es[mes_hasta][:3]} ({anio_ant} vs {anio_act})", color='#281FD0', fontweight='bold')
        
        plt.tight_layout()
        path = "temp_grafica_comparativa.png"
        plt.savefig(path)
        plt.close()
        return path

    def generar(self, resultados, output_pdf="reporte_observatorio.pdf"):
        """Compone el PDF con los resultados del procesamiento."""
        log.info(f"Generando reporte PDF: {output_pdf}")
        hoy = datetime.now()
        anio_act = hoy.year
        anio_ant = anio_act - 1
        
        # Determinar mes de corte basado en los datos procesados
        meses_detectados = []
        for r in resultados.values():
            if 'data' in r and 'MES' in r['data'].columns:
                m = r['data'][r['data']['ANIO'] == anio_act]['MES'].max()
                if pd.notnull(m): meses_detectados.append(int(m))
        mes_corte = int(max(meses_detectados)) if meses_detectados else hoy.month

        # Preparar datos de resumen
        resumen_tabla = []
        resumen_grafica = {}
        
        for nombre_delito, r in resultados.items():
            if "error" in r: continue
            df = r['data']
            ant = df[(df['ANIO'] == anio_ant) & (df['MES'] <= mes_corte)]['VALOR_NORMALIZADO'].sum()
            act = df[(df['ANIO'] == anio_act) & (df['MES'] <= mes_corte)]['VALOR_NORMALIZADO'].sum()
            
            ult_reg = "N/A"
            if 'FECHA_DT' in df.columns:
                max_d = df['FECHA_DT'].max()
                if pd.notnull(max_d):
                    ult_reg = max_d.strftime("%d/%m/%Y")
            
            # Variación
            if ant > 0:
                var = ((act - ant) / ant) * 100
                var_str = f"{var:+.1f}%"
                estado = "SUBE" if var > 0 else ("BAJA" if var < 0 else "IGUAL")
            else:
                var_str = "N/A"
                estado = "NUEVO" if act > 0 else "SIN DATOS"

            resumen_grafica[nombre_delito] = {"actual": act, "anterior": ant}
            resumen_tabla.append({
                "delito": nombre_delito,
                "anterior": int(ant),
                "actual": int(act),
                "variacion": var_str,
                "estado": estado,
                "ultimo_registro": ult_reg
            })

        # Composición del documento
        doc = SimpleDocTemplate(output_pdf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
        elements = []
        styles = getSampleStyleSheet()
        
        # Estilos personalizados
        style_title = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=18, textColor=self.azul, alignment=TA_LEFT, spaceAfter=2)
        style_subtitle = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, textColor=colors.grey, spaceAfter=10)
        style_header_table = ParagraphStyle('HT', parent=styles['Normal'], fontSize=9, textColor=colors.white, alignment=TA_CENTER, fontName='Helvetica-Bold')

        # Encabezado
        logo_path = Path("escudo_jamundi.png")
        header_table_data = [
            [
                Image(str(logo_path), 1.5*cm, 2*cm) if logo_path.exists() else "",
                [Paragraph("ALCALDÍA DE JAMUNDÍ", style_title), Paragraph(f"Observatorio del Delito - Boletín MinDefensa {self.meses_es[mes_corte]} {anio_act}", style_subtitle)],
                [Paragraph(hoy.strftime("%d/%m/%Y"), ParagraphStyle('d', alignment=TA_RIGHT, fontSize=8)), Paragraph("Confidencial - Uso Institucional", ParagraphStyle('c', alignment=TA_RIGHT, fontSize=7, textColor=colors.red))]
            ]
        ]
        header_table = Table(header_table_data, colWidths=[2*cm, 12*cm, 4*cm])
        header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        elements.append(header_table)
        elements.append(HRFlowable(width="100%", thickness=2, color=self.amarillo, spaceBefore=4, spaceAfter=10))

        # Tabla Resumen
        elements.append(Paragraph("Resumen de Indicadores Clave", styles['Heading2']))
        table_data = [[Paragraph("Indicador", style_header_table), Paragraph(str(anio_ant), style_header_table), Paragraph(str(anio_act), style_header_table), Paragraph("Variación", style_header_table), Paragraph("Estado", style_header_table), Paragraph("Últ. Reg.", style_header_table)]]
        
        for row in sorted(resumen_tabla, key=lambda x: x['actual'], reverse=True):
            color_stat = colors.red if row['estado'] == "SUBE" else (colors.green if row['estado'] == "BAJA" else colors.black)
            table_data.append([
                Paragraph(row['delito'], styles['Normal']),
                Paragraph(f"{row['anterior']:,}", ParagraphStyle('n', alignment=TA_CENTER)),
                Paragraph(f"<b>{row['actual']:,}</b>", ParagraphStyle('n', alignment=TA_CENTER)),
                Paragraph(f"<b>{row['variacion']}</b>", ParagraphStyle('n', alignment=TA_CENTER, textColor=color_stat)),
                Paragraph(row['estado'], ParagraphStyle('n', alignment=TA_CENTER, textColor=color_stat, fontSize=8)),
                Paragraph(row.get('ultimo_registro', 'N/A'), ParagraphStyle('n', alignment=TA_CENTER, fontSize=8))
            ])

        main_table = Table(table_data, colWidths=[5.5*cm, 2*cm, 2*cm, 2.5*cm, 2.5*cm, 2.5*cm])
        main_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), self.azul),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWPADDING', (0,0), (-1,-1), 6),
        ]))
        elements.append(main_table)
        elements.append(Spacer(1, 1*cm))

        # Gráfica
        graf_path = self._crear_grafica_comparativa(resumen_grafica, anio_act, anio_ant, mes_corte)
        if graf_path:
            elements.append(Paragraph("Comparativo Visual de Incidencia", styles['Heading2']))
            elements.append(Image(graf_path, 16*cm, 8*cm))
            elements.append(Spacer(1, 1*cm))

        # Firma
        elements.append(Spacer(1, 1*cm))
        style_firma = ParagraphStyle('Firma', parent=styles['Normal'], fontSize=10, textColor=colors.black, alignment=TA_CENTER)
        elements.append(Paragraph("<b>Elaborado por:</b>", style_firma))
        elements.append(Paragraph("César Alfonso Forero Molano", style_firma))
        elements.append(Paragraph("Profesional Secretaría de Seguridad y Convivencia", style_firma))

        # Pie de página institucional
        elements.append(Spacer(1, 1*cm))
        footer = Table([[Paragraph("Fuente: Ministerio de Defensa Nacional | Generado por Sistema de Vigilancia Automatizado SISC", ParagraphStyle('f', fontSize=7, alignment=TA_CENTER, textColor=colors.grey))]], colWidths=[18*cm])
        elements.append(footer)

        doc.build(elements)
        log.info("PDF generado correctamente.")
        
        # Guardar resumen en JSON para otros módulos
        import json
        with open("resumen_actual.json", "w", encoding="utf-8") as f:
            output_data = {
                "mes_corte": self.meses_es[mes_corte],
                "anio_act": anio_act,
                "indicadores": resumen_tabla
            }
            json.dump(output_data, f, ensure_ascii=False, indent=2)

        return output_pdf
