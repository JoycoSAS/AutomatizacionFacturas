@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ============================================================
REM JOYCO - FACTURAS PROCESADOR
REM Cierre trimestral real local/VPS + alertas
REM ============================================================

cd /d "%~dp0"

chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

if not exist "logs" mkdir "logs"
if not exist "logs\cierres_trimestrales_real" mkdir "logs\cierres_trimestrales_real"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i

set LOG=logs\cierres_trimestrales_real\cierre_trimestral_real_!TS!.log

echo ============================================================>> "!LOG!"
echo INICIO CIERRE TRIMESTRAL REAL CON ALERTAS - %DATE% %TIME%>> "!LOG!"
echo ============================================================>> "!LOG!"
echo Proyecto: %CD%>> "!LOG!"
echo Script: scripts\cierre_trimestral_facturas.py --real --confirmar CERRAR_TRIMESTRE>> "!LOG!"
echo Advertencia: este wrapper ejecuta el modo real, pero el script bloquea si no se cumple la fecha del trimestre.>> "!LOG!"
echo.>> "!LOG!"

python scripts\ejecutar_con_alertas.py ^
  --proceso CIERRE_TRIMESTRAL_LOCAL ^
  --asunto "Cierre trimestral real local" ^
  --mensaje-ok "Cierre trimestral real local ejecutado correctamente." ^
  --mensaje-error "Fallo el cierre trimestral real local o fue bloqueado por validacion." ^
  -- python scripts\cierre_trimestral_facturas.py --real --confirmar CERRAR_TRIMESTRE >> "!LOG!" 2>&1

set EXITCODE=%ERRORLEVEL%

echo.>> "!LOG!"
echo ============================================================>> "!LOG!"
echo FIN CIERRE TRIMESTRAL REAL CON ALERTAS - ExitCode: !EXITCODE! - %DATE% %TIME%>> "!LOG!"
echo ============================================================>> "!LOG!"

type "!LOG!"

exit /b !EXITCODE!
