import datetime
import os
import subprocess
import psutil  
from services import correo_service, zip_service, factura_service, excel_service
from config import CARPETA_ADJUNTOS, CARPETA_EXTRAIDOS

def lanzar_outlook_si_no_esta_abierto():
    # Verificar si Outlook ya está en ejecución
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'OUTLOOK.EXE' in proc.info['name'].upper():
            print("✅ Outlook ya está en ejecución.")
            return

    # Si no está, intentar iniciarlo oculto
    try:
        subprocess.Popen(["outlook.exe", "/hide"])
        print("📤 Outlook se inició en segundo plano.")
    except Exception as e:
        print(f"⚠️ No se pudo iniciar Outlook oculto: {e}")

def ejecutar_proceso():
    lanzar_outlook_si_no_esta_abierto()

    ahora = datetime.datetime.now()
    fecha, hora = ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S")

    print("\n🔍 Buscando correos recientes con adjuntos ZIP válidos...")
    correos = correo_service.obtener_correos_factura()
    if not correos:
        print("🔍 No se encontraron correos nuevos con ZIPs válidos.")
        return

    print(f"\n📥 Guardando adjuntos ZIP de {len(correos)} correos...")
    correo_service.guardar_adjuntos_zip(correos, CARPETA_ADJUNTOS)

    print("\n🗂️ Extrayendo archivos ZIP...")
    resultados = zip_service.extraer_por_zip(CARPETA_ADJUNTOS, CARPETA_EXTRAIDOS)

    historial = []
    for zipfn, carpeta in resultados:
        ruta = os.path.join(CARPETA_EXTRAIDOS, carpeta)
        regs, errores_zip = factura_service.procesar_xml_en_carpeta(ruta)

        nuevos = excel_service.guardar_en_excel(regs) if regs else 0

        if nuevos > 0 or errores_zip > 0:
            historial.append({
                'Fecha': fecha, 'Hora': hora,
                'Archivo ZIP': zipfn,
                'Nuevos XML guardados': nuevos,
                'Errores encontrados': errores_zip
            })

    if historial:
        excel_service.registrar_historial_por_zip(historial)

    print("✅ Proceso completado correctamente.")
