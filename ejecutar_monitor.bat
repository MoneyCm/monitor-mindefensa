@echo off
echo ======================================================
echo   EJECUTANDO MONITOR INTEGRAL Y SINCRONIZACION SISC
echo ======================================================
echo Fecha: %date% %time%

:: 1. Monitor Mindefensa
cd /d C:\Proyectos\monitor-mindefensa
echo [1/3] Ejecutando Monitor Mindefensa...
py monitor_ok.py >> monitor_log.txt 2>&1

:: 2. Monitor Policia
cd /d C:\Proyectos\monitor-policia
echo [2/3] Ejecutando Monitor Policia...
py descargar_policia.py >> monitor_log.txt 2>&1

:: 3. Sincronizacion con SISC Web (Neon)
cd /d C:\Proyectos\SISC-Jamundi-PRO
echo [3/3] Sincronizando datos con SISC PRO...
py sincronizar_sisc.py >> sincronizacion_log.txt 2>&1

echo ======================================================
echo   PROCESO COMPLETADO EXITOSAMENTE
echo ======================================================
pause
