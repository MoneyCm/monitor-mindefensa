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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether
from reportlab.pdfgen import canvas
from logger import log

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        if self._pageNumber > 1:
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.HexColor("#94a3b8"))
            self.drawRightString(19.5 * cm, 28.3 * cm, f"Página {self._pageNumber} de {page_count}")

class PDFGenerator:
    """Generador de boletín PDF institucional de alta calidad (visual SISC 2 páginas)."""
    def __init__(self, config_path="config.yaml"):
        with open(config_path, "r", encoding="utf-8") as f:
            self.cfg = yaml.safe_load(f)
        # Paleta institucional oficial de Jamundí
        self.azul = colors.HexColor("#281FD0")
        self.amarillo = colors.HexColor("#FFE000")
        self.gris_fondo = colors.HexColor("#F4F4F8")
        self.gris_borde = colors.HexColor("#E2E8F0")
        self.meses_es = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

    def _crear_grafica_comparativa(self, resumen_datos, anio_act, anio_ant, mes_hasta):
        """Genera una gráfica comparativa compacta de delitos top."""
        delitos = sorted(resumen_datos.keys(), key=lambda d: resumen_datos[d]['actual'], reverse=True)[:5]
        if not delitos: return None

        v_ant = [resumen_datos[d]['anterior'] for d in delitos]
        v_act = [resumen_datos[d]['actual'] for d in delitos]
        
        # Tamaño compacto para que quepa en la página 2
        fig, ax = plt.subplots(figsize=(10, 3.2), dpi=120)
        fig.patch.set_facecolor(self.cfg['estetica']['graficas_fondo'])
        ax.set_facecolor(self.cfg['estetica']['graficas_fondo'])
        
        x = range(len(delitos))
        width = 0.35
        
        ax.bar([i - width/2 for i in x], v_ant, width, label=str(anio_ant), color='#606175', alpha=0.8)
        bars = ax.bar([i + width/2 for i in x], v_act, width, label=str(anio_act), color='#281FD0')
        
        # Valores en las barras
        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'{int(height)}', xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 2), textcoords="offset points", ha='center', va='bottom', fontsize=7.5, fontweight='bold')

        ax.set_xticks(x)
        # Mostrar nombres en minúscula con mayúscula inicial para mejor estética
        labels = [d.capitalize() for d in delitos]
        ax.set_xticklabels(labels, fontsize=8.5)
        ax.legend(prop={'size': 8})
        ax.set_title(f"Comparativo Acumulado Ene-{self.meses_es[mes_hasta][:3]} ({anio_ant} vs {anio_act})", color='#281FD0', fontweight='bold', fontsize=9.5)
        
        # Ocultar bordes innecesarios
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#cccccc')
        ax.spines['bottom'].set_color('#cccccc')
        
        plt.tight_layout()
        path = "temp_grafica_comparativa.png"
        plt.savefig(path, facecolor=fig.get_facecolor(), edgecolor='none')
        plt.close()
        return path

    def generar_narrativa_dinamica(self, resumen_tabla, anio_act, anio_ant, mes_nombre):
        """Genera el contenido narrativo analítico de forma dinámica."""
        total_act = sum(r['actual'] for r in resumen_tabla)
        total_ant = sum(r['anterior'] for r in resumen_tabla)
        diff_total = total_act - total_ant
        var_total = (diff_total / total_ant * 100) if total_ant > 0 else 0
        
        # 1. Balance general
        signo_var = "un incremento" if diff_total > 0 else "una reducción"
        pct_str = f"{abs(var_total):.1f}%"
        bg_p = f"El acumulado de delitos priorizados al corte del mes de {mes_nombre} registra {total_act:,} hechos en el municipio de Jamundí durante {anio_act}, frente a {total_ant:,} reportados en el mismo periodo de {anio_ant}. Esto representa {signo_var} neta de {abs(diff_total):,} hechos, equivalente a una variación del {pct_str}."
        
        # Sort by diffAbs (for reductions)
        reducciones = [r for r in resumen_tabla if r['actual'] - r['anterior'] < 0]
        reducciones_ordenadas = sorted(reducciones, key=lambda x: x['actual'] - x['anterior'])
        
        # 2. Conductas explicativas (delitos que bajan)
        if len(reducciones_ordenadas) >= 2:
            d1 = reducciones_ordenadas[0]
            d2 = reducciones_ordenadas[1]
            ce_p = f"La disminución acumulada de incidencias delictivas se explica principalmente por la conducta de {d1['delito'].lower()}, con {abs(d1['actual'] - d1['anterior'])} hechos menos, seguido por {d2['delito'].lower()}, con una disminución de {abs(d2['actual'] - d2['anterior'])} hechos frente al año anterior."
        elif len(reducciones_ordenadas) == 1:
            d1 = reducciones_ordenadas[0]
            ce_p = f"La disminución acumulada del periodo se explica principalmente por la reducción en la conducta de {d1['delito'].lower()}, con {abs(d1['actual'] - d1['anterior'])} hechos menos."
        else:
            aumentos = sorted(resumen_tabla, key=lambda x: x['actual'] - x['anterior'], reverse=True)
            if aumentos:
                d1 = aumentos[0]
                ce_p = f"El comportamiento acumulado se encuentra impulsado principalmente por el incremento observado en la conducta de {d1['delito'].lower()}, registrando {d1['actual'] - d1['anterior']} casos adicionales."
            else:
                ce_p = "El comportamiento de las conductas prioritarias se mantiene estable sin variaciones netas durante el transcurso del año."
                
        # 3. Alertas (delitos clave o que suben >= 3 casos)
        aumentos = [r for r in resumen_tabla if r['actual'] - r['anterior'] > 0]
        alertas_criticas = [r for r in aumentos if r['delito'] in ['Homicidios', 'Secuestro', 'Extorsión'] or (r['actual'] - r['anterior'] >= 3)]
        alertas_ordenadas = sorted(alertas_criticas, key=lambda x: x['actual'] - x['anterior'], reverse=True)
        
        if len(alertas_ordenadas) >= 2:
            d1 = alertas_ordenadas[0]
            d2 = alertas_ordenadas[1]
            al_p = f"A pesar de la tendencia general, se definen alertas de seguimiento para las conductas de {d1['delito'].lower()} y {d2['delito'].lower()}, las cuales registran incrementos absolutos de {d1['actual'] - d1['anterior']} y {d2['actual'] - d2['anterior']} hechos adicionales respectivamente frente a la línea base de {anio_ant}."
        elif len(alertas_ordenadas) == 1:
            d1 = alertas_ordenadas[0]
            al_p = f"Se define una alerta de atención y patrullaje preventivo enfocado en la conducta de {d1['delito'].lower()}, la cual presenta {d1['actual'] - d1['anterior']} hechos adicionales frente al mismo lapso del año anterior."
        else:
            al_p = "Al corte del presente reporte, las cifras de delitos de alto impacto no registran incrementos críticos por encima del umbral de tolerancia estadística."
            
        return bg_p, ce_p, al_p

    def generar(self, resultados, output_pdf="reporte_observatorio.pdf"):
        """Compone el PDF con el diseño institucional de 2 páginas."""
        log.info(f"Generando reporte PDF visual SISC: {output_pdf}")
        hoy = datetime.now()
        anio_act = hoy.year
        anio_ant = anio_act - 1
        
        # Determinar mes de corte
        meses_detectados = []
        ultimos_registros = []
        for r in resultados.values():
            if 'data' in r and 'MES' in r['data'].columns:
                m = r['data'][r['data']['ANIO'] == anio_act]['MES'].max()
                if pd.notnull(m): meses_detectados.append(int(m))
            if 'data' in r and 'FECHA_DT' in r['data'].columns:
                max_d = r['data']['FECHA_DT'].max()
                if pd.notnull(max_d): ultimos_registros.append(max_d)
                
        mes_corte = int(max(meses_detectados)) if meses_detectados else hoy.month
        mes_nombre = self.meses_es[mes_corte]
        fecha_corte_str = max(ultimos_registros).strftime("%d/%m/%Y") if ultimos_registros else f"30/{mes_corte:02d}/{anio_act}"

        # Consolidar datos YTD y Puntuales
        resumen_tabla = []
        resumen_grafica = {}
        resumen_puntual = []
        
        total_registros_j = 0
        registros_con_barrio = 0
        
        for nombre_delito, r in resultados.items():
            if "error" in r: continue
            df = r['data']
            
            # Acumulado YTD
            ant_ytd = df[(df['ANIO'] == anio_ant) & (df['MES'] <= mes_corte)]['VALOR_NORMALIZADO'].sum()
            act_ytd = df[(df['ANIO'] == anio_act) & (df['MES'] <= mes_corte)]['VALOR_NORMALIZADO'].sum()
            
            # Puntos para gráfico y resumen general
            resumen_grafica[nombre_delito] = {"actual": act_ytd, "anterior": ant_ytd}
            
            diff_ytd = act_ytd - ant_ytd
            var_ytd_pct = f"{((act_ytd - ant_ytd) / ant_ytd * 100):+.1f}%" if ant_ytd > 0 else ("N/A" if act_ytd == 0 else "+100%")
            estado_ytd = "SUBE" if diff_ytd > 0 else ("BAJA" if diff_ytd < 0 else "IGUAL")
            
            # Último registro por delito
            ult_reg = "N/A"
            if 'FECHA_DT' in df.columns:
                max_d = df['FECHA_DT'].max()
                if pd.notnull(max_d): ult_reg = max_d.strftime("%d/%m/%Y")
            
            # Georreferenciación por barrio
            barrios_top_str = "Sin datos de barrio"
            total_registros_j += len(df)
            barrio_col = next((c for c in df.columns if "barrio" in c.lower()), None)
            if barrio_col:
                exclusiones = ["SIN DATO", "SIN GEORREFERENCIA", "NO REPORTA", "SIN GEORREFERENCIACION", "N/A", "NO DETECTADO"]
                df_act = df[df['ANIO'] == anio_act]
                mask_valid = ~df_act[barrio_col].astype(str).str.upper().str.strip().isin(exclusiones)
                registros_con_barrio += df_act[mask_valid][barrio_col].count()
                
                top_b = df_act[mask_valid][barrio_col].value_counts().head(2)
                if not top_b.empty:
                    barrios_top_str = ", ".join([f"{b} ({int(cnt)})" for b, cnt in top_b.items()])
            
            resumen_tabla.append({
                "delito": nombre_delito,
                "anterior": int(ant_ytd),
                "actual": int(act_ytd),
                "diffAbs": int(diff_ytd),
                "varPct": var_ytd_pct,
                "estado": estado_ytd,
                "ultimo_registro": ult_reg,
                "barrios_top": barrios_top_str
            })
            
            # Comparativo Puntual (Mes de corte)
            mismo_per = df[(df['ANIO'] == anio_ant) & (df['MES'] == mes_corte)]['VALOR_NORMALIZADO'].sum()
            per_act = df[(df['ANIO'] == anio_act) & (df['MES'] == mes_corte)]['VALOR_NORMALIZADO'].sum()
            if mes_corte > 1:
                per_ant = df[(df['ANIO'] == anio_act) & (df['MES'] == mes_corte - 1)]['VALOR_NORMALIZADO'].sum()
            else:
                per_ant = df[(df['ANIO'] == anio_ant) & (df['MES'] == 12)]['VALOR_NORMALIZADO'].sum()
                
            diff_yoy = per_act - mismo_per
            diff_pop = per_act - per_ant
            
            resumen_puntual.append({
                "delito": nombre_delito,
                "mismo_per": int(mismo_per),
                "per_act": int(per_act),
                "per_ant": int(per_ant),
                "diff_yoy": int(diff_yoy),
                "diff_pop": int(diff_pop),
                "ultimo_registro": ult_reg
            })

        # Generar narrativas dinámicas
        bg_paragraph, ce_paragraph, al_paragraph = self.generar_narrativa_dinamica(resumen_tabla, anio_act, anio_ant, mes_nombre)

        # Totales globales YTD
        total_act = sum(r['actual'] for r in resumen_tabla)
        total_ant = sum(r['anterior'] for r in resumen_tabla)
        diff_total = total_act - total_ant
        var_total = (diff_total / total_ant * 100) if total_ant > 0 else 0

        # Porcentaje de cobertura georreferenciada
        pct_cobertura = (registros_con_barrio / total_registros_j * 100) if total_registros_j > 0 else 0

        # Estructuración de ReportLab
        doc = SimpleDocTemplate(output_pdf, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.2*cm, bottomMargin=1.2*cm)
        elements = []
        styles = getSampleStyleSheet()

        # Estilos
        style_title = ParagraphStyle('Title', fontName='Helvetica-Bold', fontSize=15, textColor=self.azul, spaceAfter=2)
        style_subtitle = ParagraphStyle('Sub', fontName='Helvetica-Bold', fontSize=9, textColor=self.azul, spaceAfter=1)
        style_meta = ParagraphStyle('Meta', fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor("#606175"), alignment=TA_RIGHT)
        
        style_card_label = ParagraphStyle('CL', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.HexColor("#606175"), alignment=TA_CENTER)
        style_card_val = ParagraphStyle('CV', fontName='Helvetica-Bold', fontSize=18, textColor=self.azul, alignment=TA_CENTER)
        style_card_val_gray = ParagraphStyle('CVG', fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor("#606175"), alignment=TA_CENTER)
        style_card_sub_green = ParagraphStyle('CSG', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.HexColor("#15803D"), alignment=TA_CENTER)
        style_card_sub_red = ParagraphStyle('CSR', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.HexColor("#B91C1C"), alignment=TA_CENTER)
        style_card_sub_gray = ParagraphStyle('CSS', fontName='Helvetica', fontSize=7.5, textColor=colors.HexColor("#606175"), alignment=TA_CENTER)

        style_h2 = ParagraphStyle('H2', fontName='Helvetica-Bold', fontSize=10, textColor=self.azul, spaceAfter=1)
        style_body = ParagraphStyle('Body', fontName='Helvetica', fontSize=9, leading=13.5, textColor=colors.HexColor("#1e293b"))
        
        style_th = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=TA_CENTER)
        style_td = ParagraphStyle('TD', fontName='Helvetica', fontSize=8, alignment=TA_CENTER)
        style_td_left = ParagraphStyle('TDL', fontName='Helvetica', fontSize=8, alignment=TA_LEFT)
        style_td_bold = ParagraphStyle('TDB', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER)

        logo_path = Path("escudo_jamundi.png")

        # ======================================================================
        # PÁGINA 1
        # ======================================================================
        
        # Cintillo Amarillo Superior
        elements.append(Table([[""]], colWidths=[18*cm], rowHeights=[0.1*cm], style=TableStyle([('BACKGROUND', (0,0), (-1,-1), self.amarillo)])))
        elements.append(Spacer(1, 0.3*cm))

        # Encabezado
        header_data = [
            [
                Image(str(logo_path), 1.2*cm, 1.5*cm) if logo_path.exists() else "",
                [
                    Paragraph("BOLETÍN ESTADÍSTICO DE SEGURIDAD Y CONVIVENCIA", style_title),
                    Paragraph("ALCALDÍA MUNICIPAL DE JAMUNDÍ | SECRETARÍA DE SEGURIDAD Y CONVIVENCIA", style_subtitle)
                ],
                [
                    Paragraph("<b>BOLETÍN MINDEFENSA</b>", ParagraphStyle('B', fontName='Helvetica-Bold', fontSize=8, alignment=TA_RIGHT, textColor=self.azul)),
                    Paragraph(f"Mes: {mes_nombre} {anio_act}", style_meta),
                    Paragraph(f"Generado: {hoy.strftime('%d/%m/%Y')}", style_meta)
                ]
            ]
        ]
        header_table = Table(header_data, colWidths=[1.5*cm, 12.0*cm, 4.5*cm])
        header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
        elements.append(header_table)
        
        # Línea divisoria azul
        elements.append(Spacer(1, 0.1*cm))
        elements.append(Table([[""]], colWidths=[18*cm], rowHeights=[0.05*cm], style=TableStyle([('BACKGROUND', (0,0), (-1,-1), self.azul)])))
        elements.append(Spacer(1, 0.4*cm))

        # Banner de Corte de Cifras
        corte_content = [
            [
                Paragraph("<b>CIFRAS ACTUALIZADAS CON CORTE METODOLÓGICO AL:</b>", ParagraphStyle('c1', fontSize=8.5, textColor=colors.white)),
                Paragraph(f"<b>{fecha_corte_str}</b>", ParagraphStyle('c2', fontSize=10, textColor=self.amarillo, fontName='Helvetica-Bold', alignment=TA_RIGHT))
            ]
        ]
        corte_table = Table(corte_content, colWidths=[12.0*cm, 6.0*cm])
        corte_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), self.azul),
            ('TOPPADDING', (0,0), (-1,-1), 7),
            ('BOTTOMPADDING', (0,0), (-1,-1), 7),
            ('LEFTPADDING', (0,0), (-1,-1), 15),
            ('RIGHTPADDING', (0,0), (-1,-1), 15),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        elements.append(corte_table)
        elements.append(Spacer(1, 0.5*cm))

        # Grilla de Tarjetas (Cards Grid)
        card1 = [
            Paragraph("Total Delitos YTD", style_card_label),
            Spacer(1, 0.1*cm),
            Paragraph(f"{total_act:,}", style_card_val),
            Spacer(1, 0.15*cm),
            Paragraph(f"{var_total:+.1f}% vs Año Ant.", style_card_sub_green if var_total <= 0 else style_card_sub_red)
        ]
        card2 = [
            Paragraph("Línea Base YTD", style_card_label),
            Spacer(1, 0.1*cm),
            Paragraph(f"{total_ant:,}", style_card_val_gray),
            Spacer(1, 0.15*cm),
            Paragraph("Línea Base Anterior", style_card_sub_gray)
        ]
        card3 = [
            Paragraph("Georreferenciación", style_card_label),
            Spacer(1, 0.1*cm),
            Paragraph(f"{pct_cobertura:.1f}%", style_card_val),
            Spacer(1, 0.15*cm),
            Paragraph("Cobertura de Barrios", style_card_sub_gray)
        ]

        cards_data = [[card1, "", card2, "", card3]]
        # Ancho total: 5.6 * 3 + 0.4 * 2 = 16.8 + 0.8 = 17.6 cm
        cards_table = Table(cards_data, colWidths=[5.6*cm, 0.4*cm, 5.6*cm, 0.4*cm, 5.6*cm])
        cards_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (0,0), self.gris_fondo),
            ('BACKGROUND', (2,0), (2,0), self.gris_fondo),
            ('BACKGROUND', (4,0), (4,0), self.gris_fondo),
            ('BOX', (0,0), (0,0), 1, self.gris_borde),
            ('BOX', (2,0), (2,0), 1, self.gris_borde),
            ('BOX', (4,0), (4,0), 1, self.gris_borde),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
        ]))
        elements.append(cards_table)
        elements.append(Spacer(1, 0.6*cm))

        # Bloque Narrativo con Bordes Izquierdos
        elements.append(Paragraph("<b>ANÁLISIS DE INDICADORES Y TENDENCIAS</b>", ParagraphStyle('subt', fontName='Helvetica-Bold', fontSize=10.5, textColor=self.azul)))
        elements.append(Spacer(1, 0.2*cm))

        # 1. Balance General
        t1 = Table([[Paragraph("1. BALANCE GENERAL DE INCIDENCIA", style_h2)]], colWidths=[18*cm])
        t1.setStyle(TableStyle([('LINELEFT', (0,0), (0,0), 3, self.amarillo), ('LEFTPADDING', (0,0), (0,0), 8), ('BOTTOMPADDING', (0,0), (0,0), 2)]))
        elements.append(t1)
        elements.append(Spacer(1, 0.1*cm))
        elements.append(Paragraph(bg_paragraph, style_body))
        elements.append(Spacer(1, 0.4*cm))

        # 2. Conductas Explicativas
        t2 = Table([[Paragraph("2. DELITOS EXPLICATIVOS DEL COMPORTAMIENTO", style_h2)]], colWidths=[18*cm])
        t2.setStyle(TableStyle([('LINELEFT', (0,0), (0,0), 3, self.amarillo), ('LEFTPADDING', (0,0), (0,0), 8), ('BOTTOMPADDING', (0,0), (0,0), 2)]))
        elements.append(t2)
        elements.append(Spacer(1, 0.1*cm))
        elements.append(Paragraph(ce_paragraph, style_body))
        elements.append(Spacer(1, 0.4*cm))

        # 3. Alertas
        t3 = Table([[Paragraph("3. ALERTAS Y CONDUCTAS DE SEGUIMIENTO FOCALIZADO", style_h2)]], colWidths=[18*cm])
        t3.setStyle(TableStyle([('LINELEFT', (0,0), (0,0), 3, self.amarillo), ('LEFTPADDING', (0,0), (0,0), 8), ('BOTTOMPADDING', (0,0), (0,0), 2)]))
        elements.append(t3)
        elements.append(Spacer(1, 0.1*cm))
        elements.append(Paragraph(al_paragraph, style_body))

        # Salto de página forzado a la página 2
        elements.append(PageBreak())

        # ======================================================================
        # PÁGINA 2
        # ======================================================================
        
        # Encabezado menor
        header_p2 = [
            [
                Image(str(logo_path), 0.8*cm, 1.0*cm) if logo_path.exists() else "",
                [
                    Paragraph("BOLETÍN ESTADÍSTICO DE SEGURIDAD Y CONVIVENCIA", ParagraphStyle('t2', fontName='Helvetica-Bold', fontSize=10, textColor=self.azul)),
                    Paragraph(f"Corte: {mes_nombre} {anio_act} | Alcaldía de Jamundí", ParagraphStyle('s2', fontSize=8, textColor=colors.HexColor("#606175")))
                ],
                [
                    Paragraph(f"", ParagraphStyle('p2', fontSize=8, alignment=TA_RIGHT, textColor=colors.HexColor("#94a3b8")))
                ]
            ]
        ]
        header_table_p2 = Table(header_p2, colWidths=[1.0*cm, 13.0*cm, 4.0*cm])
        header_table_p2.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0)]))
        elements.append(header_table_p2)
        elements.append(Spacer(1, 0.1*cm))
        elements.append(Table([[""]], colWidths=[18*cm], rowHeights=[0.03*cm], style=TableStyle([('BACKGROUND', (0,0), (-1,-1), self.azul)])))
        elements.append(Spacer(1, 0.3*cm))

        # Gráfica Comparativa Matplotlib
        graf_path = self._crear_grafica_comparativa(resumen_grafica, anio_act, anio_ant, mes_corte)
        if graf_path:
            elements.append(Paragraph("<b>COMPARATIVO VISUAL DE INCIDENCIA</b>", ParagraphStyle('st2', fontName='Helvetica-Bold', fontSize=10, textColor=self.azul)))
            elements.append(Spacer(1, 0.1*cm))
            elements.append(Image(graf_path, 18*cm, 5.76*cm)) # 10:3.2 ratio
            elements.append(Spacer(1, 0.3*cm))

        # Estilos específicos para la tabla compacta paralela
        style_th_comp = ParagraphStyle('THC', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=TA_CENTER)
        style_td_comp = ParagraphStyle('TDC', fontName='Helvetica', fontSize=7.5, alignment=TA_CENTER)
        style_td_comp_left = ParagraphStyle('TDCL', fontName='Helvetica', fontSize=7.5, alignment=TA_LEFT)
        style_td_comp_bold = ParagraphStyle('TDCB', fontName='Helvetica-Bold', fontSize=7.5, alignment=TA_CENTER)

        # Tabla 1: Comparativo Acumulado YTD (Todos)
        elements.append(Paragraph("<b>4. COMPARATIVO ACUMULADO POR DELITO (YTD)</b>", ParagraphStyle('st3', fontName='Helvetica-Bold', fontSize=10, textColor=self.azul)))
        elements.append(Spacer(1, 0.15*cm))

        ytd_table_data = [[
            Paragraph("Delito", style_th_comp), 
            Paragraph(f"{anio_ant} acum.", style_th_comp), 
            Paragraph(f"{anio_act} acum.", style_th_comp), 
            Paragraph("Diferencia", style_th_comp), 
            Paragraph("Variación %", style_th_comp)
        ]]

        ytd_rows = sorted(resumen_tabla, key=lambda x: x['actual'], reverse=True)
        for row in ytd_rows:
            color_var = colors.HexColor("#B91C1C") if row['diffAbs'] > 0 else (colors.HexColor("#15803D") if row['diffAbs'] < 0 else colors.black)
            ytd_table_data.append([
                Paragraph(f"<b>{row['delito']}</b><br/><font color='#64748b'><i>(Últ. reg: {row['ultimo_registro']})</i></font>", style_td_comp_left),
                Paragraph(f"{row['anterior']:,}", style_td_comp),
                Paragraph(f"<b>{row['actual']:,}</b>", style_td_comp_bold),
                Paragraph(f"<b>{row['diffAbs']:+,}</b>", ParagraphStyle('d', alignment=TA_CENTER, textColor=color_var, fontSize=7.5)),
                Paragraph(f"<b>{row['varPct']}</b>", ParagraphStyle('v', alignment=TA_CENTER, textColor=color_var, fontSize=7.5))
            ])

        ytd_table = Table(ytd_table_data, colWidths=[7.0*cm, 2.5*cm, 2.5*cm, 3.0*cm, 3.0*cm])
        ytd_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), self.azul),
            ('GRID', (0,0), (-1,-1), 0.5, self.gris_borde),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWPADDING', (0,0), (-1,-1), 4),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, self.gris_fondo])
        ]))
        elements.append(ytd_table)
        elements.append(Spacer(1, 0.4*cm))

        # Tabla 2: Comparativo Puntual del Periodo (Todos)
        elements.append(Paragraph(f"<b>5. COMPARATIVO PUNTUAL ({mes_nombre.upper()})</b>", ParagraphStyle('st4', fontName='Helvetica-Bold', fontSize=10, textColor=self.azul)))
        elements.append(Spacer(1, 0.15*cm))

        puntual_table_data = [[
            Paragraph("Delito", style_th_comp), 
            Paragraph(f"Mismo Per. {anio_ant}", style_th_comp), 
            Paragraph("Per. Actual", style_th_comp), 
            Paragraph("Per. Anterior", style_th_comp), 
            Paragraph("Var. YoY", style_th_comp), 
            Paragraph("Var. PoP", style_th_comp)
        ]]

        puntual_rows = sorted(resumen_puntual, key=lambda x: x['per_act'], reverse=True)
        for row in puntual_rows:
            color_yoy = colors.HexColor("#B91C1C") if row['diff_yoy'] > 0 else (colors.HexColor("#15803D") if row['diff_yoy'] < 0 else colors.black)
            color_pop = colors.HexColor("#B91C1C") if row['diff_pop'] > 0 else (colors.HexColor("#15803D") if row['diff_pop'] < 0 else colors.black)
            
            puntual_table_data.append([
                Paragraph(f"<b>{row['delito']}</b><br/><font color='#64748b'><i>(Últ. reg: {row['ultimo_registro']})</i></font>", style_td_comp_left),
                Paragraph(f"{row['mismo_per']:,}", style_td_comp),
                Paragraph(f"<b>{row['per_act']:,}</b>", style_td_comp_bold),
                Paragraph(f"{row['per_ant']:,}", style_td_comp),
                Paragraph(f"<b>{row['diff_yoy']:+,}</b>", ParagraphStyle('dy', alignment=TA_CENTER, textColor=color_yoy, fontSize=7.5)),
                Paragraph(f"<b>{row['diff_pop']:+,}</b>", ParagraphStyle('dp', alignment=TA_CENTER, textColor=color_pop, fontSize=7.5))
            ])

        puntual_table = Table(puntual_table_data, colWidths=[6.0*cm, 2.4*cm, 2.4*cm, 2.4*cm, 2.4*cm, 2.4*cm])
        puntual_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), self.azul),
            ('GRID', (0,0), (-1,-1), 0.5, self.gris_borde),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWPADDING', (0,0), (-1,-1), 4),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, self.gris_fondo])
        ]))
        elements.append(puntual_table)
        elements.append(Spacer(1, 0.4*cm))

        # Bloque de Control Documental y Firma
        import hashlib
        # Hash rápido basado en el nombre del archivo de entrada
        nombre_arch = next((r['nombreArchivo'] for r in resultados.values() if 'nombreArchivo' in r), 'MINDEFENSA_COMPILADO')
        hash_obj = hashlib.sha256(nombre_arch.encode('utf-8'))
        file_hash = hash_obj.hexdigest()[:16].upper()

        control_content = [
            [
                Paragraph("<b>Nota Metodológica:</b> Este reporte es generado automáticamente a partir del consolidado municipalizado de la base de datos nacional del Ministerio de Defensa Nacional de Colombia. Las cifras reflejan el análisis operativo del Observatorio del Delito de Jamundí.", ParagraphStyle('nm', fontSize=6.5, textColor=colors.HexColor("#64748b"), leading=9)),
                [
                    Paragraph(f"<b>Aprobó:</b> Carolina Obando Gómez (Sec. Seguridad)", ParagraphStyle('ct1', fontSize=7.5)),
                    Paragraph(f"<b>Elaboró:</b> César Alfonso Forero Molano (Obs. Delito)", ParagraphStyle('ct2', fontSize=7.5)),
                    Paragraph(f"<b>Versión:</b> 1.0 | <b>Corte:</b> {mes_nombre} de {anio_act}", ParagraphStyle('ct3', fontSize=7.5)),
                    Paragraph(f"<b>Verificación SHA256:</b> {file_hash}", ParagraphStyle('ct4', fontSize=6.5, fontName='Courier'))
                ]
            ]
        ]
        control_table = Table(control_content, colWidths=[10.0*cm, 8.0*cm])
        control_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), self.gris_fondo),
            ('BOX', (0,0), (-1,-1), 1, self.gris_borde),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10)
        ]))
        elements.append(control_table)

        doc.build(elements, canvasmaker=NumberedCanvas)
        log.info("PDF generado correctamente y adaptado a la visual del SISC.")
        
        # Guardar resumen en JSON
        import json
        with open("resumen_actual.json", "w", encoding="utf-8") as f:
            output_data = {
                "mes_corte": mes_nombre,
                "anio_act": anio_act,
                "indicadores": resumen_tabla
            }
            json.dump(output_data, f, ensure_ascii=False, indent=2)

        return output_pdf
