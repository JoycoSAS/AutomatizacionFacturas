# services/factura_service.py

import os
import re
import xml.etree.ElementTree as ET
from html import unescape

import PyPDF2  # pip install PyPDF2

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


def _extraer_descripciones_completas(root):
    """
    Extrae la descripción de cada InvoiceLine probando distintos nodos:
      1. Item/Description
      2. Item/Name
      3. InvoiceLine/Note
      4. SellersItemIdentification/ID
    Y concatena todas con “; ”.
    """
    descripciones = []
    for linea in root.findall('.//{*}InvoiceLine'):
        texto = None

        # 1) Item/Description
        nodo = linea.find('.//{*}Item/{*}Description')
        if nodo is not None and nodo.text:
            texto = nodo.text.strip()

        # 2) Item/Name
        if not texto:
            nodo = linea.find('.//{*}Item/{*}Name')
            if nodo is not None and nodo.text:
                texto = nodo.text.strip()

        # 3) InvoiceLine/Note
        if not texto:
            nodo = linea.find('{*}Note')
            if nodo is not None and nodo.text:
                texto = nodo.text.strip()

        # 4) SellersItemIdentification/ID
        if not texto:
            nodo = linea.find('.//{*}SellersItemIdentification/{*}ID')
            if nodo is not None and nodo.text:
                texto = nodo.text.strip()

        if texto:
            descripciones.append(texto)

    return "; ".join(descripciones)


def _extraer_actividad_de_pdf(xml_path):
    carpeta = os.path.dirname(xml_path)
    for fn in os.listdir(carpeta):
        if fn.lower().endswith('.pdf'):
            pdf_path = os.path.join(carpeta, fn)
            try:
                reader = PyPDF2.PdfReader(pdf_path)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
                m = re.search(r'(?:CIIU|Actividad\s+Econ[oó]mica)[^\d]*(\d{4,5})',
                              text, re.IGNORECASE)
                if m:
                    return m.group(1)
            except Exception as e:
                errores.append(f"Error leyendo PDF '{fn}': {e}")
    return ""


def leer_datos_xml(path):
    try:
        inner_xml = _extract_inner_invoice(path)
        root = ET.fromstring(inner_xml) if inner_xml else ET.parse(path).getroot()
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

    # — Datos básicos —
    emisor      = obtener_texto(root,
                    './/cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name', ns)
    cliente     = obtener_texto(root,
                    './/cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name', ns)
    if not cliente or cliente.lower() == 'no aplica':
        cliente = obtener_texto(
            root,
            './/cac:AccountingCustomerParty//cac:PartyLegalEntity/cbc:RegistrationName', ns
        )
    if not cliente:
        cliente = obtener_texto(
            root,
            './/cac:AccountingCustomerParty//cac:PartyIdentification/cbc:ID', ns
        )

    numero      = obtener_texto(root, './cbc:ID', ns)
    descripcion_lineas = _extraer_descripciones_completas(root)

    nit         = obtener_texto(
                    root,
                    './/cac:AccountingSupplierParty//cac:PartyLegalEntity/cbc:CompanyID', ns
                  )
    tipo_contribuyente = obtener_texto(root,
                    './/cac:PartyTaxScheme/cbc:TaxLevelCode', ns)
    fecha_text  = obtener_texto(root, './cbc:IssueDate', ns)
    cufe        = obtener_texto(root, './/cbc:UUID', ns)

    # — Ciudad emisora —
    ciudad_nombre = obtener_texto(
        root,
        './/cac:AccountingSupplierParty//cac:PhysicalLocation//cac:Address//cbc:CityName',
        ns
    )
    ciudad_codigo = obtener_texto(
        root,
        './/cac:AccountingSupplierParty//cac:PhysicalLocation//cac:Address//cbc:ID',
        ns
    )

    # — Totales —
    subtotal = convertir_a_numero(obtener_texto(
                   root, './/cac:LegalMonetaryTotal/cbc:LineExtensionAmount', ns))
    total    = convertir_a_numero(obtener_texto(
                   root, './/cac:LegalMonetaryTotal/cbc:PayableAmount', ns))

    # — Actividad económica —
    act_eco_el = root.find('.//{*}IndustryClassificationCode')
    actividad_economica = (
        act_eco_el.text.strip()
        if act_eco_el is not None and act_eco_el.text else ""
    )
    if not actividad_economica:
        raw_xml = ET.tostring(root, encoding='utf-8', method='xml').decode('utf-8')
        m = re.search(
            r'(?:CIIU|Actividad\s+Econ[oó]mica)[^\d]*(\d{4,5})',
            raw_xml, re.IGNORECASE
        )
        if m:
            actividad_economica = m.group(1)
    if not actividad_economica:
        actividad_economica = _extraer_actividad_de_pdf(path)

    # — IVA discriminado —
    iva_5 = iva_19 = 0.0
    for tax in root.findall('./cac:TaxTotal/cac:TaxSubtotal', ns):
        amt      = convertir_a_numero(obtener_texto(tax, './cbc:TaxAmount', ns))
        pct_text = obtener_texto(tax, './cac:TaxCategory/cbc:Percent', ns)
        try:
            pct = float(pct_text)
            if abs(pct - 5.0)  < 0.01:
                iva_5  += amt
            elif abs(pct - 19.0) < 0.01:
                iva_19 += amt
        except:
            continue

    # — Retenciones —
    reteiva = reteica = rete_fuente = 0.0
    for tax in root.findall('./cac:WithholdingTaxTotal/cac:TaxSubtotal', ns):
        amt      = convertir_a_numero(obtener_texto(tax, './cbc:TaxAmount', ns))
        tax_id   = obtener_texto(
                      tax,
                      './cac:TaxCategory/cac:TaxScheme/cbc:ID', ns
                  ).strip().lower()
        tax_name = obtener_texto(
                      tax,
                      './cac:TaxCategory/cac:TaxScheme/cbc:Name', ns
                  ).strip().lower()

        if tax_id == '05' or 'iva' in tax_name:
            reteiva    += amt
        elif tax_id == '06' or 'fuente' in tax_name or 'renta' in tax_name:
            rete_fuente += amt
        elif tax_id == '07' or 'ica' in tax_name:
            reteica    += amt

    reteiva     = -abs(reteiva)
    reteica     = -abs(reteica)
    rete_fuente = -abs(rete_fuente)
    total      += reteiva + reteica + rete_fuente

    return {
        "Archivo":                os.path.basename(path),
        "Empresa emisora":        emisor,
        "CUFE":                   cufe,
        "Ciudad emisora":         ciudad_nombre,
        "Código ciudad":          ciudad_codigo,
        "NIT":                    nit,
        "Cliente":                cliente,
        "Número de factura":      numero,
        "Año":                    fecha_text[:4],
        "Mes":                    fecha_text[5:7],
        "Día":                    fecha_text[8:10],
        "Tipo de contribuyente":  tipo_contribuyente,
        "Actividad económica":    actividad_economica,
        "DescripcionLineas":      descripcion_lineas,
        "Subtotal":               subtotal,
        "IVA 5%":                 iva_5,
        "IVA 19%":                iva_19,
        "Retención de IVA":       reteiva,
        "Retención de ICA":       reteica,
        "Retención en la fuente": rete_fuente,
        "Total":                  total
    }


def procesar_xml_en_carpeta(ruta_carpeta):
    registros   = []
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
