@echo off
cd /d "C:\Users\Infraestructura\Downloads\facturas_procesador"
python procesador_facturas.py
if %errorlevel% neq 0 (
    echo Error: El script de Python no se ejecutó correctamente.
    exit /b %errorlevel%
)