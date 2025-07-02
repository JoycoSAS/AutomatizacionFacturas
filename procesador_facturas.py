import win32com.client
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import datetime
from html import unescape

# --- CONFIGURACIÓN ---
BASE_DIR           = r"C:\Users\Infraestructura\Downloads\Facturas_PROCESADOR"
CARPETA_ADJUNTOS   = os.path.join(BASE_DIR, "adjuntos")
CARPETA_EXTRAIDOS  = os.path.join(BASE_DIR, "extraidos")
ARCHIVO_EXCEL      = os.path.join(BASE_DIR, "facturas.xlsx")
HISTORIAL_EXCEL    = os.path.join(BASE_DIR, "historial_ejecuciones.xlsx")
STORE_NAME         = "auxiliar.infraestructura@joyco.com.co"
TEMP_CHECK_DIR     = os.path.join(BASE_DIR, "temp_check")

errores = []

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

def guardar_adjuntos_zip(correos):
    os.makedirs(CARPETA_ADJUNTOS, exist_ok=True)
    count = 0
    for correo in correos:
        try:
            for adj in correo.Attachments:
                if adj.FileName.lower().endswith(".zip"):
                    safe_name = re.sub(r'[<>:"/\\|?*]', '_', adj.FileName)
                    destino = os.path.join(CARPETA_ADJUNTOS, safe_name)
                    try:
                        adj.SaveAsFile(destino)
                        count += 1
                        print(f"📅 ZIP guardado: {safe_name}")
                    except Exception as e:
                        errores.append(f"Error guardando ZIP '{safe_name}': {e}")
        except Exception as e:
            errores.append(f"Error accediendo adjuntos de correo: {e}")
    return count

def extraer_por_zip():
    os.makedirs(CARPETA_EXTRAIDOS, exist_ok=True)
    resultados = []
    for zipfn in os.listdir(CARPETA_ADJUNTOS):
        if not zipfn.lower().endswith(".zip"):
            continue
        src = os.path.join(CARPETA_ADJUNTOS, zipfn)
        carpeta = os.path.splitext(zipfn)[0]
        dst = os.path.join(CARPETA_EXTRAIDOS, carpeta)
        os.makedirs(dst, exist_ok=True)
        try:
            with zipfile.ZipFile(src, 'r') as z:
                z.extractall(dst)
            print(f"🖜️ Extraído {zipfn} en {carpeta}/")
            resultados.append((zipfn, carpeta))
        except Exception as e:
            errores.append(f"Error extrayendo ZIP '{zipfn}': {e}")
    return resultados

def _extract_inner_invoice(path):
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        desc = root.find('.//{*}Attachment//{*}ExternalReference//{*}Description')
        if desc is not None and desc.text:
            raw = unescape(desc.text.strip())
            if raw.lstrip().startswith('<'):
                return raw
    except Exception as e:
        errores.append(f"Error extrayendo XML interno en {os.path.basename(path)}: {e}")
    return None

def leer_datos_xml(path):
    try:
        inner_xml = _extract_inner_invoice(path)
        if inner_xml:
            root = ET.fromstring(inner_xml)
        else:
            root = ET.parse(path).getroot()
    except ET.ParseError as e:
        errores.append(f"XML mal formado '{path}': {e}")
        return None
    except Exception as e:
        errores.append(f"Error al leer XML '{path}': {e}")
        return None

    def obtener_texto(elem, tag, ns, default=""):
        nodo = elem.find(tag, ns)
        return nodo.text.strip() if nodo is not None and nodo.text else default

    def convertir_a_numero(texto):
        try:
            return float(re.sub(r'[^\d.]', '', texto.replace(',', '.')))
        except:
            return 0.0

    ns = {
        'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
        'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'
    }

    emisor = obtener_texto(root, './/cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name', ns)
    cliente = obtener_texto(root, './/cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name', ns)
    numero = obtener_texto(root, './cbc:ID', ns)
    descripcion = obtener_texto(root, './/cac:InvoiceLine/cac:Item/cbc:Description', ns)
    subtotal = convertir_a_numero(obtener_texto(root, './/cac:LegalMonetaryTotal/cbc:LineExtensionAmount', ns))
    total = convertir_a_numero(obtener_texto(root, './/cac:LegalMonetaryTotal/cbc:PayableAmount', ns))
    nit = obtener_texto(root, './/cac:AccountingSupplierParty/cac:Party/cac:PartyLegalEntity/cbc:CompanyID', ns)
    tipo_contribuyente = obtener_texto(root, './/cac:PartyTaxScheme/cbc:TaxLevelCode', ns)

    act_eco_el = root.find('.//{*}AccountingSupplierParty//{*}Party//{*}IndustryClassificationCode')
    actividad_economica = act_eco_el.text.strip() if act_eco_el is not None and act_eco_el.text else ""

    regimen_responsabilidad = obtener_texto(root, './/cac:PartyTaxScheme/cac:TaxScheme/cbc:Name', ns).lower()
    responsable_iva = "sí" if "iva" in regimen_responsabilidad else "no"
    fecha = obtener_texto(root, './cbc:IssueDate', ns)

    iva_5 = iva_19 = reteiva = reteica = 0.0

    for tax in root.findall('./cac:TaxTotal/cac:TaxSubtotal', ns) + root.findall('./cac:WithholdingTaxTotal/cac:TaxSubtotal', ns):
        amount = convertir_a_numero(obtener_texto(tax, './cbc:TaxAmount', ns))
        percent = obtener_texto(tax, './cac:TaxCategory/cbc:Percent', ns)
        code = obtener_texto(tax, './cac:TaxCategory/cac:TaxScheme/cbc:ID', ns)
        name = obtener_texto(tax, './cac:TaxCategory/cac:TaxScheme/cbc:Name', ns).upper()
        try:
            percent_val = float(percent)
            if abs(percent_val - 5.0) < 0.01:
                iva_5 += amount
            elif abs(percent_val - 19.0) < 0.01:
                iva_19 += amount
        except:
            pass
        if code == "06" or ("IVA" in name and not percent):
            reteiva += amount
        elif code == "07" or "ICA" in name:
            reteica += amount

    reteiva = -abs(reteiva)
    reteica = -abs(reteica)
    total += reteiva + reteica

    return {
        "Empresa emisora": emisor,
        "NIT": nit,
        "Cliente": cliente,
        "Número de factura": numero,
        "Concepto": descripcion,
        "Subtotal": subtotal,
        "IVA 5%": iva_5,
        "IVA 19%": iva_19,
        "Retención de IVA": reteiva,
        "Retención de ICA": reteica,
        "Total": total,
        "Año": fecha[:4],
        "Mes": fecha[5:7],
        "Día": fecha[8:10],
        "Tipo de contribuyente": tipo_contribuyente,
        "Responsable de IVA": responsable_iva,
        "Actividad económica": actividad_economica,
        "Archivo": os.path.basename(path)
    }

def guardar_en_excel(datos):
    columnas = [
        "Empresa emisora", "NIT", "Cliente", "Número de factura", "Concepto", "Subtotal",
        "IVA 5%", "IVA 19%", "Retención de IVA", "Retención de ICA", "Total",
        "Año", "Mes", "Día", "Tipo de contribuyente",
        "Responsable de IVA", "Actividad económica", "Archivo"
    ]
    df = pd.DataFrame(datos)[columnas]
    nuevos = 0
    if os.path.exists(ARCHIVO_EXCEL):
        viejo = pd.read_excel(ARCHIVO_EXCEL, engine='openpyxl')
        combo = pd.concat([viejo, df], ignore_index=True)
        combo = combo.drop_duplicates(subset=['Número de factura'], keep='last')
        nuevos = len(combo) - len(viejo)
        final = combo
    else:
        nuevos = len(df)
        final = df
    final.to_excel(ARCHIVO_EXCEL, index=False)
    print(f"✅ Excel actualizado: {ARCHIVO_EXCEL}")
    return nuevos

def registrar_historial_por_zip(filas):
    dfh = pd.DataFrame(filas)
    if os.path.exists(HISTORIAL_EXCEL):
        viejo = pd.read_excel(HISTORIAL_EXCEL, engine='openpyxl')
        out = pd.concat([viejo, dfh], ignore_index=True)
    else:
        out = dfh
    out.to_excel(HISTORIAL_EXCEL, index=False)
    print(f"📑 Historial actualizado: {HISTORIAL_EXCEL}")

def ejecutar_proceso():
    errores.clear()
    ahora = datetime.datetime.now()
    fecha, hora = ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S")

    correos = obtener_correos_factura()
    if not correos:
        print("🔍 No hay correos nuevos.")
        return

    guardar_adjuntos_zip(correos)
    resultados = extraer_por_zip()

    historial = []
    for zipfn, carpeta in resultados:
        regs, errs_zip = [], 0
        ruta = os.path.join(CARPETA_EXTRAIDOS, carpeta)
        for fn in os.listdir(ruta):
            if fn.lower().endswith('.xml'):
                full = os.path.join(ruta, fn)
                d = leer_datos_xml(full)
                if d:
                    regs.append(d)
                    print(f"✅ Procesado: {fn}")
                else:
                    errs_zip += 1
        nuevos = guardar_en_excel(regs) if regs else 0
        if nuevos > 0 or errs_zip > 0:
            historial.append({
                'Fecha': fecha, 'Hora': hora,
                'Archivo ZIP': zipfn,
                'Nuevos XML guardados': nuevos,
                'Errores encontrados': errs_zip
            })

    if historial:
        registrar_historial_por_zip(historial)

    if errores:
        print("\n⚠️ Se presentaron errores:")
        for e in errores:
            print(f" - {e}")

if __name__ == '__main__':
    ejecutar_proceso()
