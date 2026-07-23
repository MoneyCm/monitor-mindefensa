import os
import sys
from pathlib import Path
from procesador import DataProcessor
from generar_reporte import PDFGenerator

print("🧪 PROBANDO GENERADOR DE REPORTES PDF SISC...")
print("="*60)

# Cargar archivos locales del monitor
processor = DataProcessor()
archivos = {
    "Homicidios": "HOMICIDIO INTENCIONAL.xlsx",
    "Secuestro": "SECUESTRO.xlsx",
    "Hurtos a Personas": "HURTO PERSONAS.xlsx",
    "Violencia Intrafamiliar": "VIOLENCIA INTRAFAMILIAR.xlsx"
}

resultados = {}
for delito, arch in archivos.items():
    ruta = Path(arch)
    if not ruta.exists():
        print(f"⚠️ Archivo {arch} no encontrado localmente. Saltando.")
        continue
    
    print(f"⚙️ Procesando {delito} ({arch})...")
    res = processor.procesar_archivo(ruta)
    if res and "error" not in res:
        resultados[delito] = res
        print(f"   Conteo Jamundí: {res['total_jamundi']} casos.")
    else:
        print(f"   ❌ Error ETL: {res.get('error') if res else 'Desconocido'}")

if not resultados:
    print("❌ No se pudieron procesar datos. Deteniendo.")
    sys.exit(1)

print("\n🎨 Generando PDF de prueba...")
pdf_gen = PDFGenerator()
pdf_path = pdf_gen.generar(resultados, output_pdf="reporte_observatorio_prueba.pdf")
print(f"✅ ¡PDF Generado con éxito!: {pdf_path}")
print("="*60)
