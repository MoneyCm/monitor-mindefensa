import os
import sys
import yaml
from pathlib import Path
from procesador import DataProcessor
from generar_reporte import PDFGenerator

print("🧪 PROBANDO GENERADOR DE REPORTES PDF SISC (DELITOS RECIENTES)...")
print("="*60)

# Cargar Configuración
with open("config.yaml", "r", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

processor = DataProcessor()
resultados = {}

# Carpeta de salida del scraper
xlsx_dir = Path("mindefensa_xlsx")

# Detectar y procesar todos los delitos configurados
for d in cfg['datasets']:
    nombre_dataset = d['nombre']
    patron = d['patron'].upper()
    
    # Buscar archivo local que coincida con el patrón en mindefensa_xlsx
    archivos_locales = list(xlsx_dir.glob("*.xlsx"))
    archivo_encontrado = next((a for a in archivos_locales if patron in a.name.upper() and not a.name.startswith("~$")), None)
    
    if archivo_encontrado:
        print(f"⚙️ Procesando {nombre_dataset} ({archivo_encontrado.name})...")
        res = processor.procesar_archivo(archivo_encontrado)
        if res and "error" not in res:
            resultados[nombre_dataset] = res
            print(f"   Conteo Jamundí: {res['total_jamundi']} casos.")
        else:
            print(f"   ❌ Error ETL en {archivo_encontrado.name}: {res.get('error') if res else 'Desconocido'}")
    else:
        print(f"⚠️ No se encontró archivo local para el patrón: {patron}")

if not resultados:
    print("❌ No se pudieron procesar datos. Deteniendo.")
    sys.exit(1)

print("\n🎨 Generando PDF de prueba...")
pdf_gen = PDFGenerator()
pdf_path = pdf_gen.generar(resultados, output_pdf="reporte_observatorio_prueba.pdf")
print(f"✅ ¡PDF Generado con éxito!: {pdf_path}")
print("="*60)
