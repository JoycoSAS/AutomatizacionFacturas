# services/zip_service.py

import os
import zipfile
from utils.logger import errores


def extraer_por_zip(carpeta_adjuntos, carpeta_extraidos):
    """
    Extrae los archivos contenidos en los ZIP ubicados en la carpeta de adjuntos.
    Retorna una lista con tuplas (nombre_zip, carpeta_destino).
    """
    os.makedirs(carpeta_extraidos, exist_ok=True)
    resultados = []

    for zipfn in os.listdir(carpeta_adjuntos):
        if not zipfn.lower().endswith(".zip"):
            continue

        src = os.path.join(carpeta_adjuntos, zipfn)
        carpeta = os.path.splitext(zipfn)[0]
        dst = os.path.join(carpeta_extraidos, carpeta)
        os.makedirs(dst, exist_ok=True)

        try:
            with zipfile.ZipFile(src, 'r') as z:
                z.extractall(dst)
            print(f"🗜️ Extraído {zipfn} en {carpeta}/")
            resultados.append((zipfn, carpeta))
        except Exception as e:
            errores.append(f"Error extrayendo ZIP '{zipfn}': {e}")

    return resultados
