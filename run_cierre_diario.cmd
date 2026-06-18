@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM JOYCO - FACTURAS PROCESADOR
REM Cierre diario local + subida SharePoint + alertas
REM ============================================================

chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

cd /d "%~dp0"

if not exist "logs" mkdir "logs"
if not exist "logs\cierres" mkdir "logs\cierres"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i

set LOG_FILE=logs\cierres\cierre_diario_!TS!.log

echo ============================================================ >> "!LOG_FILE!"
echo INICIO CIERRE DIARIO LOCAL CON ALERTAS - !TS! >> "!LOG_FILE!"
echo Carpeta: %CD% >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"

python scripts\ejecutar_con_alertas.py ^
  --proceso CIERRE_DIARIO_LOCAL ^
  --asunto "Cierre diario local" ^
  --mensaje-error "Fallo la generacion local del cierre diario. No se ejecuta subida a SharePoint." ^
  -- python scripts\cierre_diario_seguro.py >> "!LOG_FILE!" 2>&1

set EXIT_CIERRE=%ERRORLEVEL%

echo ------------------------------------------------------------ >> "!LOG_FILE!"
echo FIN CIERRE LOCAL - ExitCode: !EXIT_CIERRE! >> "!LOG_FILE!"
echo ------------------------------------------------------------ >> "!LOG_FILE!"

if not "!EXIT_CIERRE!"=="0" (
    echo ERROR: cierre local fallo. No se ejecuta subida SharePoint. >> "!LOG_FILE!"
    exit /b !EXIT_CIERRE!
)

echo ============================================================ >> "!LOG_FILE!"
echo INICIO SUBIDA CIERRE DIARIO SHAREPOINT CON ALERTAS >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"

python scripts\ejecutar_con_alertas.py ^
  --proceso CIERRE_DIARIO_SHAREPOINT ^
  --asunto "Subida cierre diario a SharePoint" ^
  --mensaje-error "Fallo la subida del cierre diario a SharePoint." ^
  -- python scripts\subir_cierre_diario_sharepoint.py >> "!LOG_FILE!" 2>&1

set EXIT_UPLOAD=%ERRORLEVEL%

echo ------------------------------------------------------------ >> "!LOG_FILE!"
echo FIN SUBIDA SHAREPOINT - ExitCode: !EXIT_UPLOAD! >> "!LOG_FILE!"
echo ------------------------------------------------------------ >> "!LOG_FILE!"

exit /b !EXIT_UPLOAD!
