# services/factura_service.py

import os
import re
import base64
import xml.etree.ElementTree as ET
from html import unescape

import PyPDF2  # pip install PyPDF2

from utils.helpers import obtener_texto, convertir_a_numero
from utils.logger import errores


# --------------------------------------------------------------------
# Fallback opcional: sólo para "Actividad económica" (CIIU) si el XML
# no trae el dato. No se usa para totales ni para crear filas mínimas.
# --------------------------------------------------------------------
PDF_FALLBACK_ENABLED = True

# ----------------------------
# Helpers de robustez XML
# ----------------------------
_CTRL_REGEX = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")  # quita controles ilegales XML 1.0
_AMP_FIX = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+|#x[0-9A-Fa-f]+);)")  # & que no inicia entidad

def _clean_xml_text(txt: str) -> str:
    """Limpia controles y ampersands sueltos para que ET pueda parsear."""
    txt = _CTRL_REGEX.sub("", txt)
    txt = _AMP_FIX.sub("&amp;", txt)
    return txt

def _safe_parse_xml(path: str) -> ET.Element:
    """
    Parse tolerante:
      - Lee en binario y decodifica con utf-8-sig (soporta BOM)
      - Reintenta tras limpieza si ET.fromstring falla
    """
    with open(path, "rb") as f:
        raw = f.read()
    try:
        text = raw.decode("utf-8-sig", errors="replace")
    except Exception:
        text = raw.decode(errors="replace")

    try:
        return ET.fromstring(text)
    except ET.ParseError:
        text2 = _clean_xml_text(text)
        return ET.fromstring(text2)


def _extract_inner_invoice(path: str) -> str | None:
    """
    Devuelve el XML del Invoice embebido si existe en un AttachedDocument.
    Soporta:
      - <EmbeddedDocumentBinaryObject> (base64 con el XML)
      - <ExternalReference>/<Description> (XML escapado)
      - <ExternalReference>/<URI> apuntando a un XML vecino
    """
    try:
        # parse tolerante también para el "envoltorio"
        root = _safe_parse_xml(path)

        # 1) Binario base64
        bin_el = root.find('.//{*}EmbeddedDocumentBinaryObject')
        if bin_el is not None and bin_el.text:
            try:
                raw = base64.b64decode(bin_el.text.strip(), validate=False)
                txt = raw.decode('utf-8', errors='ignore').lstrip()
                if txt.startswith('<'):
                    return txt
            except Exception as e:
                errores.append(
                    f"EmbeddedDocumentBinaryObject inválido en {os.path.basename(path)}: {e}"
                )

        # 2) Description con XML escapado
        desc = root.find('.//{*}Attachment//{*}ExternalReference//{*}Description')
        if desc is not None and desc.text:
            raw = unescape(desc.text.strip())
            if raw.lstrip().startswith('<'):
                return raw

        # 3) URI a un archivo local
        uri = root.find('.//{*}Attachment//{*}ExternalReference//{*}URI')
        if uri is not None and uri.text:
            maybe = uri.text.strip()
            # Si no parece URL absoluta, probar como archivo vecino
            if not re.match(r'^[a-z]+://', maybe, flags=re.I):
                carpeta = os.path.dirname(path)
                destino = os.path.join(carpeta, os.path.basename(maybe))
                if os.path.exists(destino) and destino.lower().endswith('.xml'):
                    try:
                        with open(destino, 'r', encoding='utf-8', errors='ignore') as f:
                            txt = f.read().lstrip()
                            if txt.startswith('<'):
                                return txt
                    except Exception as e:
                        errores.append(
                            f"No se pudo abrir URI '{maybe}' en {os.path.basename(path)}: {e}"
                        )

    except Exception as e:
        errores.append(f"Error extrayendo XML interno en {os.path.basename(path)}: {e}")

    return None


def _extraer_descripciones_completas(root: ET.Element) -> str:
    """
    Extrae la descripción de cada InvoiceLine probando:
      1. Item/Description
      2. Item/Name
      3. InvoiceLine/Note
      4. SellersItemIdentification/ID
    Concatena con “; ”.
    """
    descripciones = []
    for linea in root.findall('.//{*}InvoiceLine'):
        texto = None

        nodo = linea.find('.//{*}Item/{*}Description')
        if nodo is not None and nodo.text:
            texto = nodo.text.strip()

        if not texto:
            nodo = linea.find('.//{*}Item/{*}Name')
            if nodo is not None and nodo.text:
                texto = nodo.text.strip()

        if not texto:
            nodo = linea.find('{*}Note')
            if nodo is not None and nodo.text:
                texto = nodo.text.strip()

        if not texto:
            nodo = linea.find('.//{*}SellersItemIdentification/{*}ID')
            if nodo is not None and nodo.text:
                texto = nodo.text.strip()

        if texto:
            descripciones.append(texto)

    return "; ".join(descripciones)


def _extraer_actividad_de_pdf(xml_path: str) -> str:
    """
    ÚLTIMO recurso (opcional) para detectar CIIU desde el PDF
    vecino. No se usa para ningún otro dato.
    """
    if not PDF_FALLBACK_ENABLED:
        return ""
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


def leer_datos_xml(path: str) -> dict | None:
    """
    Lee una factura UBL. Si es AttachedDocument intenta extraer el Invoice
    embebido. Si no hay Invoice embebido, NO crea fila (evita “fila mínima”).
    Ahora es tolerante a XML con caracteres ilegales o & sin escapar.
    """
    try:
        inner_xml = _extract_inner_invoice(path)
        if inner_xml:
            # El XML interno también puede venir sucio: limpiamos si es necesario
            try:
                root = ET.fromstring(inner_xml)
            except ET.ParseError:
                root = ET.fromstring(_clean_xml_text(inner_xml))
        else:
            # Parse tolerante del archivo principal
            root = _safe_parse_xml(path)
            # Si es AttachedDocument y no conseguimos Invoice interno -> nada
            local = root.tag.split('}')[-1] if '}' in root.tag else root.tag
            if local == 'AttachedDocument':
                errores.append(
                    f"AttachedDocument sin Invoice embebido: {os.path.basename(path)}"
                )
                return None
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
    emisor = obtener_texto(
        root, './/cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name', ns
    )

    cliente = obtener_texto(
        root, './/cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name', ns
    )
    if not cliente or cliente.lower() == 'no aplica':
        cliente = obtener_texto(
            root,
            './/cac:AccountingCustomerParty//cac:PartyLegalEntity/cbc:RegistrationName',
            ns
        )
    if not cliente:
        cliente = obtener_texto(
            root, './/cac:AccountingCustomerParty//cac:PartyIdentification/cbc:ID', ns
        )

    numero = obtener_texto(root, './cbc:ID', ns)
    descripcion_lineas = _extraer_descripciones_completas(root)

    nit = obtener_texto(
        root, './/cac:AccountingSupplierParty//cac:PartyLegalEntity/cbc:CompanyID', ns
    )
    tipo_contribuyente = obtener_texto(
        root, './/cac:PartyTaxScheme/cbc:TaxLevelCode', ns
    )
    fecha_text = obtener_texto(root, './cbc:IssueDate', ns)
    cufe = obtener_texto(root, './/cbc:UUID', ns)

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
    subtotal = convertir_a_numero(
        obtener_texto(root, './/cac:LegalMonetaryTotal/cbc:LineExtensionAmount', ns)
    )
    total = convertir_a_numero(
        obtener_texto(root, './/cac:LegalMonetaryTotal/cbc:PayableAmount', ns)
    )

    # — Actividad económica —
    act_eco_el = root.find('.//{*}IndustryClassificationCode')
    actividad_economica = (
        act_eco_el.text.strip()
        if act_eco_el is not None and act_eco_el.text else ""
    )
    if not actividad_economica:
        raw_xml = ET.tostring(root, encoding='utf-8', method='xml').decode('utf-8', errors='ignore')
        m = re.search(r'(?:CIIU|Actividad\s+Econ[oó]mica)[^\d]*(\d{4,5})',
                      raw_xml, re.IGNORECASE)
        if m:
            actividad_economica = m.group(1)
    if not actividad_economica:
        actividad_economica = _extraer_actividad_de_pdf(path)

    # — IVA discriminado —
    iva_5 = iva_19 = 0.0
    for tax in root.findall('./cac:TaxTotal/cac:TaxSubtotal', ns):
        amt = convertir_a_numero(obtener_texto(tax, './cbc:TaxAmount', ns))
        pct_text = obtener_texto(tax, './cac:TaxCategory/cbc:Percent', ns)
        try:
            pct = float(pct_text)
            if abs(pct - 5.0) < 0.01:
                iva_5 += amt
            elif abs(pct - 19.0) < 0.01:
                iva_19 += amt
        except Exception:
            continue

    # — Retenciones —
    reteiva = reteica = rete_fuente = 0.0
    for tax in root.findall('./cac:WithholdingTaxTotal/cac:TaxSubtotal', ns):
        amt = convertir_a_numero(obtener_texto(tax, './cbc:TaxAmount', ns))
        tax_id = obtener_texto(
            tax, './cac:TaxCategory/cac:TaxScheme/cbc:ID', ns
        ).strip().lower()
        tax_name = obtener_texto(
            tax, './cac:TaxCategory/cac:TaxScheme/cbc:Name', ns
        ).strip().lower()

        if tax_id == '05' or 'iva' in tax_name:
            reteiva += amt
        elif tax_id == '06' or 'fuente' in tax_name or 'renta' in tax_name:
            rete_fuente += amt
        elif tax_id == '07' or 'ica' in tax_name:
            reteica += amt

    # Guardar retenciones negativas y ajustar total neto
    reteiva = -abs(reteiva)
    reteica = -abs(reteica)
    rete_fuente = -abs(rete_fuente)
    total += reteiva + reteica + rete_fuente

    return {
        "Archivo":                os.path.basename(path),
        "Empresa emisora":        emisor,
        "CUFE":                   cufe,
        "Ciudad emisora":         ciudad_nombre,
        "Código ciudad":          ciudad_codigo,
        "NIT":                    nit,
        "Cliente":                cliente,
        "Número de factura":      numero,
        "Año":                    (fecha_text or "")[:4],
        "Mes":                    (fecha_text or "")[5:7],
        "Día":                    (fecha_text or "")[8:10],
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


def procesar_xml_en_carpeta(ruta_carpeta: str) -> tuple[list[dict], int]:
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
