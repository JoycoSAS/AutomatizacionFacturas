import os
import xml.etree.ElementTree as ET
import re
from html import unescape

from utils.helpers import obtener_texto, convertir_a_numero
from utils.logger import errores


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


def procesar_xml_en_carpeta(ruta_carpeta):
    registros = []
    errores_zip = 0

    for archivo in os.listdir(ruta_carpeta):
        if archivo.lower().endswith('.xml'):
            full_path = os.path.join(ruta_carpeta, archivo)
            datos = leer_datos_xml(full_path)
            if datos:
                registros.append(datos)
                print(f"✅ Procesado: {archivo}")
            else:
                errores_zip += 1

    return registros, errores_zip
