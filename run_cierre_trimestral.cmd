@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

if not exist "logs\cierres_trimestrales" (
    mkdir "logs\cierres_trimestrales"
)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set TS=%%i

set LOG=logs\cierres_trimestrales\cierre_trimestral_%TS%.log

echo ============================================================>> "%LOG%"
echo INICIO CIERRE TRIMESTRAL LOCAL - %DATE% %TIME%>> "%LOG%"
echo ============================================================>> "%LOG%"
echo Proyecto: %CD%>> "%LOG%"
echo Script: scripts\cierre_trimestral_facturas.py --local>> "%LOG%"
echo.>> "%LOG%"

python scripts\cierre_trimestral_facturas.py --local >> "%LOG%" 2>&1
set EXITCODE=%ERRORLEVEL%

echo.>> "%LOG%"
echo ============================================================>> "%LOG%"
echo FIN CIERRE TRIMESTRAL LOCAL - ExitCode: %EXITCODE% - %DATE% %TIME%>> "%LOG%"
echo ============================================================>> "%LOG%"

type "%LOG%"

exit /b %EXITCODE%
