# services/correo_service.py

import os
import re
import zipfile
import win32com.client
from config import STORE_NAME, TEMP_CHECK_DIR
from utils.logger import errores


def obtener_correos_factura():
    """
    Obtiene correos recientes que contienen al menos un archivo .zip
    que contenga al menos un archivo .xml válido.
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mi_store = outlook.Folders[STORE_NAME]

        try:
            bandeja = mi_store.Folders["Inbox"]
        except:
            bandeja = mi_store.Folders["Bandeja de entrada"]

        items = bandeja.Items
        items.Sort("[ReceivedTime]", True)

        correos_validos = []
        os.makedirs(TEMP_CHECK_DIR, exist_ok=True)

        for correo in items:
            try:
                for adj in correo.Attachments:
                    if adj.FileName.lower().endswith(".zip"):
                        temp_path = os.path.join(TEMP_CHECK_DIR, adj.FileName)
                        adj.SaveAsFile(temp_path)
                        with zipfile.ZipFile(temp_path, 'r') as z:
                            if any(f.filename.lower().endswith(".xml") for f in z.infolist()):
                                correos_validos.append(correo)
                                break
                        os.remove(temp_path)
            except Exception as e:
                errores.append(f"Error leyendo ZIP de correo: {e}")
                continue

        return correos_validos

    except Exception as e:
        errores.append(f"Error accediendo a Outlook: {e}")
        return []



def guardar_adjuntos_zip(correos, carpeta_adjuntos):
    """
    Guarda los archivos .zip adjuntos de los correos en la carpeta especificada.
    """
    os.makedirs(carpeta_adjuntos, exist_ok=True)
    count = 0
    for correo in correos:
        try:
            for adj in correo.Attachments:
                if adj.FileName.lower().endswith(".zip"):
                    safe_name = re.sub(r'[<>:"/\\|?*]', '_', adj.FileName)
                    destino = os.path.join(carpeta_adjuntos, safe_name)
                    try:
                        adj.SaveAsFile(destino)
                        count += 1
                        print(f"📥 ZIP guardado: {safe_name}")
                    except Exception as e:
                        errores.append(f"Error guardando ZIP '{safe_name}': {e}")
        except Exception as e:
            errores.append(f"Error accediendo adjuntos de correo: {e}")
    return count
