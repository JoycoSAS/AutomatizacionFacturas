@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM JOYCO - FACTURAS PROCESADOR
REM Ejecucion produccion local controlada
REM ============================================================

chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

cd /d "%~dp0"

if not exist "logs" mkdir "logs"
if not exist "logs\produccion" mkdir "logs\produccion"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i

set LOG_FILE=logs\produccion\facturas_produccion_!TS!.log

echo ============================================================ >> "!LOG_FILE!"
echo INICIO FACTURAS PRODUCCION LOCAL - !TS! >> "!LOG_FILE!"
echo Carpeta: %CD% >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"

python main_aprobadas_integrado.py >> "!LOG_FILE!" 2>&1

set EXIT_CODE=%ERRORLEVEL%

echo ============================================================ >> "!LOG_FILE!"
echo FIN FACTURAS PRODUCCION LOCAL - ExitCode: !EXIT_CODE! >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"

exit /b !EXIT_CODE!
