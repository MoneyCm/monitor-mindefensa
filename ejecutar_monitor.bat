@echo off
echo ======================================================
echo   EJECUTANDO MONITOR INTEGRAL Y SINCRONIZACION SISC
echo ======================================================
echo Fecha: %date% %time%

:: 1. Monitor Mindefensa + Ingesta SISC
cd /d C:\Proyectos\monitor-mindefensa
echo [1/3] Ejecutando Monitor Mindefensa + Ingesta SISC...
py monitor_sisc.py >> monitor_log.txt 2>&1

:: 2. Monitor Policia
cd /d C:\Proyectos\monitor-policia
echo [2/3] Ejecutando Monitor Policia...
py descargar_policia.py >> monitor_log.txt 2>&1

:: 3. Fin del Proceso
echo [3/3] Monitores completados.
echo ======================================================
echo   PROCESO COMPLETADO EXITOSAMENTE
echo ======================================================
pause
