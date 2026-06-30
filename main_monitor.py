import os
import sys
import yaml
import requests
from pathlib import Path
from datetime import datetime
from logger import log
from scraper import MinDefensaScraper
from estado import StateManager
from procesador import DataProcessor
from generar_reporte import PDFGenerator
from exportar_sisc import SISCExporter
from notificar import Notifier

def main():
    log.info("=" * 60)
    log.info("MONITOR MINDEFENSA V2.0 - SISTEMA PROFESIONAL")
    log.info("=" * 60)

    # 1. Cargar Configuración
    with open("config.yaml", "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    # 2. Scrape e Intercepción
    scraper = MinDefensaScraper()
    archivos_detectados = scraper.ejecutar()
    
    if not archivos_detectados:
        log.error("No se detectaron archivos en el portal. Abortando.")
        sys.exit(0) # Salida limpia pero sin cambios

    # 3. Comparar Estado
    state = StateManager()
    nuevos, cambiados = state.comparar(archivos_detectados)
    
    log.info(f"Reporte de Cambios: {len(nuevos)} nuevos, {len(cambiados)} actualizados.")

    # 4. Determinar si necesitamos descargar y procesar
    # Forzamos descarga en días de reporte o si hay cambios
    dia_semana = datetime.now().weekday() # 1=Martes, 4=Viernes
    tipo_run = "normal"
    if dia_semana == 1: tipo_run = "reunion"
    if dia_semana == 4: tipo_run = "consejo"
    
    force_all = os.environ.get("FORCE_DOWNLOAD", "false").lower() == "true" or tipo_run != "normal"
    
    para_descargar = []
    if force_all:
        log.info(f"Modo FORZADO ({tipo_run}): Procesando toda la lista detectada.")
        para_descargar = archivos_detectados
    else:
        para_descargar = nuevos + cambiados

    if not para_descargar:
        log.info("Sin novedades y no es día de reporte. Finalizando.")
        state.save(archivos_detectados, 0, 0)
        # Escribir output de GitHub para saltar pasos
        if 'GITHUB_OUTPUT' in os.environ:
            with open(os.environ['GITHUB_OUTPUT'], 'a') as f:
                f.write("hay_cambios=false\n")
        return

    # 5. Descarga y Procesamiento
    processor = DataProcessor()
    resultados = {}
    
    # Filtrar solo archivos de interés basados en patrones de config
    filtros = [d['patron'].upper() for d in cfg['datasets']]
    
    for a in para_descargar:
        nombre = a['nombre'].upper()
        dataset_info = next((d for d in cfg['datasets'] if d['patron'].upper() in nombre), None)
        
        if not dataset_info:
            continue # Ignorar archivos que no nos interesan
            
        log.info(f"Descargando: {a['nombre']}")
        try:
            r = requests.get(a['url'], timeout=60)
            if r.status_code == 200:
                ruta_local = processor.output_dir / a['nombre']
                ruta_local.write_bytes(r.content)
                
                # Procesar ETL
                log.info(f"Procesando ETL: {a['nombre']}")
                res = processor.procesar_archivo(ruta_local)
                if res and "error" not in res:
                    resultados[dataset_info['nombre']] = res
                    log.info(f"  Result: {res['total_jamundi']} casos Jamundí detectados.")
                else:
                    log.warning(f"  Error ETL: {res.get('error') if res else 'Desconocido'}")
        except Exception as e:
            log.error(f"Error procesando {a['nombre']}: {e}")

    # 6. Guardar Estado
    state.save(archivos_detectados, len(nuevos), len(cambiados))

    # 6.5. Cargar archivos locales para datasets no descargados en esta ejecución
    # Esto garantiza que el reporte contenga todos los delitos configurados
    log.info("Completando reporte con archivos locales para delitos sin cambios...")
    for d in cfg['datasets']:
        nombre_dataset = d['nombre']
        if nombre_dataset not in resultados:
            # Buscar el archivo correspondiente en la lista detectada
            patron = d['patron'].upper()
            archivo_info = next((a for a in archivos_detectados if patron in a['nombre'].upper()), None)
            if archivo_info:
                ruta_local = processor.output_dir / archivo_info['nombre']
                if ruta_local.exists():
                    log.info(f"Cargando {nombre_dataset} desde caché local: {archivo_info['nombre']}")
                    try:
                        res = processor.procesar_archivo(ruta_local)
                        if res and "error" not in res:
                            resultados[nombre_dataset] = res
                        else:
                            log.warning(f"  Error al procesar archivo local {archivo_info['nombre']}: {res.get('error') if res else 'Desconocido'}")
                    except Exception as e:
                        log.error(f"Error procesando archivo local {archivo_info['nombre']}: {e}")
                else:
                    log.warning(f"  Archivo local para {nombre_dataset} no encontrado en {ruta_local}")
            else:
                log.warning(f"  No se detectó ningún archivo en el portal para el patrón: {d['patron']}")

    # 7. Generar Reporte PDF
    reporte_ok = False
    if resultados:
        pdf_gen = PDFGenerator()
        pdf_path = pdf_gen.generar(resultados)
        reporte_ok = True
        
        # 8. Exportar a SISC API
        exporter = SISCExporter()
        exporter.exportar(resultados)
        
        # 9. Notificar
        notifier = Notifier()
        notifier.enviar(pdf_path, tipo_run=tipo_run)

    # Output GitHub Actions
    if 'GITHUB_OUTPUT' in os.environ:
        with open(os.environ['GITHUB_OUTPUT'], 'a') as f:
            f.write(f"hay_cambios=true\n")
            f.write(f"reporte_generado={str(reporte_ok).lower()}\n")

    log.info("PIPELINE COMPLETADO EXITOSAMENTE")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.critical(f"FALLO CATASTRÓFICO: {e}")
        import traceback
        log.error(traceback.format_exc())
        sys.exit(1)
