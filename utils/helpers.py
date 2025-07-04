import re
import xml.etree.ElementTree as ET
from html import unescape


def convertir_a_numero(texto):
    try:
        return float(re.sub(r'[^\d.]', '', texto.replace(',', '.')))
    except:
        return 0.0


def obtener_texto(elem, tag, ns, default=""):
    nodo = elem.find(tag, ns)
    return nodo.text.strip() if nodo is not None and nodo.text else default


def obtener_actividad_economica(root):
    act_eco_el = root.find('.//{*}AccountingSupplierParty//{*}Party//{*}IndustryClassificationCode')
    return act_eco_el.text.strip() if act_eco_el is not None and act_eco_el.text else ""


def extraer_inner_invoice(path):
    """
    Extrae el XML interno si está incrustado dentro de la etiqueta Description.
    """
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        desc = root.find('.//{*}Attachment//{*}ExternalReference//{*}Description')
        if desc is not None and desc.text:
            raw = unescape(desc.text.strip())
            if raw.lstrip().startswith('<'):
                return raw
    except Exception as e:
        return None
    return None
