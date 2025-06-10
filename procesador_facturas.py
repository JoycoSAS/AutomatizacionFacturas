import win32com.client
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import datetime

# ────── CONFIGURACIÓN ──────────────
BASE_DIR           = r"C:\Users\Infraestructura\Downloads\Facturas_PROCESADOR"
CARPETA_ADJUNTOS   = os.path.join(BASE_DIR, "adjuntos")
CARPETA_EXTRAIDOS  = os.path.join(BASE_DIR, "extraidos")
ARCHIVO_EXCEL      = os.path.join(BASE_DIR, "facturas.xlsx")
HISTORIAL_EXCEL    = os.path.join(BASE_DIR, "historial_ejecuciones.xlsx")
STORE_NAME         = "auxiliar.infraestructura@joyco.com.co"
# ───────────────────────────────────

def obtener_correos_factura():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mi_store = outlook.Folders[STORE_NAME]
        try:
            bandeja = mi_store.Folders["Inbox"]
        except:
            bandeja = mi_store.Folders["Bandeja de entrada"]
        items = bandeja.Items
        items.Sort("[ReceivedTime]", True)
        return [c for c in items if "factura" in (c.Subject or "").lower()]
    except Exception as e:
        errores.append(f"Error obteniendo correos: {e}")
        return []

def guardar_adjuntos_zip(correos):
    os.makedirs(CARPETA_ADJUNTOS, exist_ok=True)
    count = 0
    for correo in correos:
        try:
            for adj in correo.Attachments:
                if adj.FileName.lower().endswith(".zip"):
                    fn = re.sub(r'[<>:"/\\\\|?*]', '_', adj.FileName)
                    dest = os.path.join(CARPETA_ADJUNTOS, fn)
                    try:
                        adj.SaveAsFile(dest)
                        count += 1
                        print(f"📥 ZIP guardado: {fn}")
                    except Exception as e:
                        errores.append(f"Error guardando ZIP '{fn}': {e}")
        except Exception as e:
            errores.append(f"Error accediendo adjuntos de correo: {e}")
    return count

def extraer_por_zip():
    os.makedirs(CARPETA_EXTRAIDOS, exist_ok=True)
    resultados = []
    for zipfn in os.listdir(CARPETA_ADJUNTOS):
        if not zipfn.lower().endswith(".zip"):
            continue
        zip_path = os.path.join(CARPETA_ADJUNTOS, zipfn)
        folder_name = os.path.splitext(zipfn)[0]
        dst = os.path.join(CARPETA_EXTRAIDOS, folder_name)
        os.makedirs(dst, exist_ok=True)
        try:
            with zipfile.ZipFile(zip_path, 'r') as z:
                z.extractall(dst)
            print(f"🗜️ Extraído {zipfn} en {folder_name}/")
            resultados.append((zipfn, folder_name))
        except Exception as e:
            errores.append(f"Error extrayendo ZIP '{zipfn}': {e}")
    return resultados

def _extract_cdata_inner(path):
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        desc = root.find('.//{*}Description')
        if desc is not None and desc.text and desc.text.strip().startswith('<'):
            return desc.text.strip()
    except Exception as e:
        errores.append(f"Error extrayendo CDATA en {os.path.basename(path)}: {e}")
    return None

def leer_datos_xml(path):
    try:
        raw = _extract_cdata_inner(path)
        if raw:
            inner = ET.fromstring(raw)
        else:
            inner = ET.parse(path).getroot()
    except ET.ParseError as e:
        errores.append(f"XML mal formado '{path}': {e}")
        return None
    except Exception as e:
        errores.append(f"Error general leyendo XML '{path}': {e}")
        return None

    def find_local(elem, name):
        for e in elem.iter():
            if e.tag.split('}')[-1] == name:
                return e
        return None

    uuid_el   = find_local(inner, 'UUID')
    issue_el  = find_local(inner, 'IssueDate')
    time_el   = find_local(inner, 'IssueTime')
    total_el  = find_local(inner, 'PayableAmount') or find_local(inner, 'LineExtensionAmount')
    emisor_el = find_local(inner, 'RegistrationName')
    receptor_el = None
    for e in inner.findall('.//{*}ReceiverParty//{*}PartyTaxScheme//{*}RegistrationName'):
        receptor_el = e
        break

    return {
        'Fuente':   os.path.basename(path),
        'UUID':     uuid_el.text if uuid_el is not None else '',
        'Fecha':    issue_el.text if issue_el is not None else '',
        'Hora':     time_el.text if time_el is not None else '',
        'Total':    total_el.text if total_el is not None else '',
        'Emisor':   emisor_el.text if emisor_el is not None else '',
        'Receptor': receptor_el.text if receptor_el is not None else ''
    }

def guardar_en_excel(datos):
    cols = ['Fuente','UUID','Fecha','Hora','Total','Emisor','Receptor']
    df = pd.DataFrame(datos)[cols]
    nuevos = 0
    if os.path.exists(ARCHIVO_EXCEL):
        viejo = pd.read_excel(ARCHIVO_EXCEL)
        combinados = pd.concat([viejo, df])
        combinados = combinados.drop_duplicates(subset=['UUID'], keep='last')
        nuevos = len(combinados) - len(viejo)
        df_final = combinados
    else:
        nuevos = len(df)
        df_final = df
    df_final.to_excel(ARCHIVO_EXCEL, index=False)
    print(f"✅ Excel actualizado: {ARCHIVO_EXCEL}")
    return nuevos

def registrar_historial_por_zip(filas):
    df_nuevo = pd.DataFrame(filas)
    if os.path.exists(HISTORIAL_EXCEL):
        df_viejo = pd.read_excel(HISTORIAL_EXCEL)
        df = pd.concat([df_viejo, df_nuevo], ignore_index=True)
    else:
        df = df_nuevo
    df.to_excel(HISTORIAL_EXCEL, index=False)
    print(f"📑 Historial actualizado: {HISTORIAL_EXCEL}")

def ejecutar_proceso():
    global errores
    errores = []

    ahora = datetime.datetime.now()
    fecha = ahora.strftime("%Y-%m-%d")
    hora = ahora.strftime("%H:%M:%S")

    correos = obtener_correos_factura()
    if not correos:
        print("🔍 No hay correos nuevos.")
        return

    guardar_adjuntos_zip(correos)
    zip_resultados = extraer_por_zip()

    historial_filas = []
    for zipfn, carpeta in zip_resultados:
        registros = []
        errores_zip = 0
        d = os.path.join(CARPETA_EXTRAIDOS, carpeta)
        for fn in os.listdir(d):
            if fn.lower().endswith('.xml'):
                path = os.path.join(d, fn)
                data = leer_datos_xml(path)
                if data:
                    registros.append(data)
                    print(f"✅ Procesado: {fn}")
                else:
                    errores_zip += 1

        nuevos = guardar_en_excel(registros) if registros else 0
        if nuevos > 0 or errores_zip > 0:
            historial_filas.append({
                'Fecha': fecha,
                'Hora': hora,
                'Archivo ZIP': zipfn,
                'Nuevos XML guardados': nuevos,
                'Errores encontrados': errores_zip
            })

    if historial_filas:
        registrar_historial_por_zip(historial_filas)

if __name__ == '__main__':
    ejecutar_proceso()
