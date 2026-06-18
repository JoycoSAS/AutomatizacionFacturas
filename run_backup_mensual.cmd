@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM JOYCO - FACTURAS PROCESADOR
REM Backup mensual local + subida SharePoint + alertas
REM ============================================================

chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

cd /d "%~dp0"

if not exist "logs" mkdir "logs"
if not exist "logs\backups_mensuales" mkdir "logs\backups_mensuales"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i

set LOG_FILE=logs\backups_mensuales\backup_mensual_!TS!.log

echo ============================================================ >> "!LOG_FILE!"
echo INICIO BACKUP MENSUAL LOCAL CON ALERTAS - !TS! >> "!LOG_FILE!"
echo Carpeta: %CD% >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"

python scripts\ejecutar_con_alertas.py ^
  --proceso BACKUP_MENSUAL_LOCAL ^
  --asunto "Backup mensual local" ^
  --mensaje-ok "Backup mensual local generado correctamente." ^
  --mensaje-error "Fallo la generacion local del backup mensual. No se ejecuta subida a SharePoint." ^
  -- python scripts\backup_mensual_seguro.py >> "!LOG_FILE!" 2>&1

set EXIT_BACKUP=%ERRORLEVEL%

echo ------------------------------------------------------------ >> "!LOG_FILE!"
echo FIN BACKUP MENSUAL LOCAL - ExitCode: !EXIT_BACKUP! >> "!LOG_FILE!"
echo ------------------------------------------------------------ >> "!LOG_FILE!"

if not "!EXIT_BACKUP!"=="0" (
    echo ERROR: backup mensual local fallo. No se ejecuta subida SharePoint. >> "!LOG_FILE!"
    exit /b !EXIT_BACKUP!
)

echo ============================================================ >> "!LOG_FILE!"
echo INICIO SUBIDA BACKUP MENSUAL SHAREPOINT CON ALERTAS >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"

python scripts\ejecutar_con_alertas.py ^
  --proceso BACKUP_MENSUAL_SHAREPOINT ^
  --asunto "Subida backup mensual a SharePoint" ^
  --mensaje-ok "Backup mensual subido y verificado correctamente en SharePoint." ^
  --mensaje-error "Fallo la subida del backup mensual a SharePoint." ^
  -- python scripts\subir_backup_mensual_sharepoint.py >> "!LOG_FILE!" 2>&1

set EXIT_UPLOAD=%ERRORLEVEL%

echo ------------------------------------------------------------ >> "!LOG_FILE!"
echo FIN SUBIDA BACKUP MENSUAL SHAREPOINT - ExitCode: !EXIT_UPLOAD! >> "!LOG_FILE!"
echo ------------------------------------------------------------ >> "!LOG_FILE!"

exit /b !EXIT_UPLOAD!

