# services/factura_service.py
import os
import re
import base64
import xml.etree.ElementTree as ET
from html import unescape
from decimal import Decimal

import PyPDF2  # pip install PyPDF2

from utils.helpers import obtener_texto, convertir_a_numero
from utils.logger import errores


PDF_FALLBACK_ENABLED = True


_CTRL_REGEX = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_AMP_FIX = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+|#x[0-9A-Fa-f]+);)")


def _clean_xml_text(txt: str) -> str:
    txt = _CTRL_REGEX.sub("", txt or "")
    txt = _AMP_FIX.sub("&amp;", txt)
    return txt


def _safe_parse_xml(path: str) -> ET.Element:
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
    try:
        root = _safe_parse_xml(path)

        bin_el = root.find(".//{*}EmbeddedDocumentBinaryObject")
        if bin_el is not None and bin_el.text:
            try:
                raw = base64.b64decode(bin_el.text.strip(), validate=False)
                txt = raw.decode("utf-8", errors="ignore").lstrip()
                if txt.startswith("<"):
                    return txt
            except Exception as e:
                errores.append(f"EmbeddedDocumentBinaryObject inválido en {os.path.basename(path)}: {e}")

        desc = root.find(".//{*}Attachment//{*}ExternalReference//{*}Description")
        if desc is not None and desc.text:
            raw = unescape(desc.text.strip())
            if raw.lstrip().startswith("<"):
                return raw

        uri = root.find(".//{*}Attachment//{*}ExternalReference//{*}URI")
        if uri is not None and uri.text:
            maybe = uri.text.strip()
            if not re.match(r"^[a-z]+://", maybe, flags=re.I):
                carpeta = os.path.dirname(path)
                destino = os.path.join(carpeta, os.path.basename(maybe))
                if os.path.exists(destino) and destino.lower().endswith(".xml"):
                    try:
                        with open(destino, "r", encoding="utf-8", errors="ignore") as f:
                            txt = f.read().lstrip()
                            if txt.startswith("<"):
                                return txt
                    except Exception as e:
                        errores.append(f"No se pudo abrir URI '{maybe}' en {os.path.basename(path)}: {e}")

    except Exception as e:
        errores.append(f"Error extrayendo XML interno en {os.path.basename(path)}: {e}")

    return None


def _local_name(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def _iter_lineas_documento(root: ET.Element) -> list[ET.Element]:
    for xp in (".//{*}InvoiceLine", ".//{*}CreditNoteLine", ".//{*}DebitNoteLine"):
        lines = root.findall(xp)
        if lines:
            return lines
    return []


def _linea_descripcion(linea: ET.Element) -> str:
    nodo = linea.find(".//{*}Item/{*}Description")
    if nodo is not None and nodo.text and nodo.text.strip():
        return nodo.text.strip()

    nodo = linea.find(".//{*}Item/{*}Name")
    if nodo is not None and nodo.text and nodo.text.strip():
        return nodo.text.strip()

    for n in linea.findall("{*}Note"):
        if n is not None and n.text and n.text.strip():
            return n.text.strip()

    nodo = linea.find(".//{*}SellersItemIdentification/{*}ID")
    if nodo is not None and nodo.text and nodo.text.strip():
        return nodo.text.strip()

    return ""


def _linea_iva_percent(linea: ET.Element) -> float | None:
    pct = linea.find(".//{*}TaxTotal/{*}TaxSubtotal/{*}TaxCategory/{*}Percent")
    if pct is not None and pct.text:
        try:
            return float(str(pct.text).strip())
        except Exception:
            return None

    pct2 = linea.find(".//{*}ClassifiedTaxCategory/{*}Percent")
    if pct2 is not None and pct2.text:
        try:
            return float(str(pct2.text).strip())
        except Exception:
            return None

    return None


def _extraer_descripciones_por_iva(root: ET.Element) -> dict:
    all_items: list[str] = []
    iva19: list[str] = []
    iva5: list[str] = []
    iva0: list[str] = []

    for linea in _iter_lineas_documento(root):
        desc = _linea_descripcion(linea)
        if not desc:
            continue

        if desc not in all_items:
            all_items.append(desc)

        pct = _linea_iva_percent(linea)
        if pct is None or abs(pct - 0.0) < 0.0001:
            if desc not in iva0:
                iva0.append(desc)
        elif abs(pct - 5.0) < 0.01:
            if desc not in iva5:
                iva5.append(desc)
        elif abs(pct - 19.0) < 0.01:
            if desc not in iva19:
                iva19.append(desc)
        else:
            if desc not in iva0:
                iva0.append(desc)

    return {
        "all": "; ".join(all_items),
        "iva19": "; ".join(iva19),
        "iva5": "; ".join(iva5),
        "iva0": "; ".join(iva0),
    }


def _extraer_actividad_de_pdf(xml_path: str) -> str:
    if not PDF_FALLBACK_ENABLED:
        return ""
    carpeta = os.path.dirname(xml_path)
    for fn in os.listdir(carpeta):
        if fn.lower().endswith(".pdf"):
            pdf_path = os.path.join(carpeta, fn)
            try:
                reader = PyPDF2.PdfReader(pdf_path)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
                m = re.search(r"(?:CIIU|Actividad\s+Econ[oó]mica)[^\d]*(\d{4,5})", text, re.IGNORECASE)
                if m:
                    return m.group(1)
            except Exception as e:
                errores.append(f"Error leyendo PDF '{fn}': {e}")
    return ""


def _dec_from_str(s: str | None) -> Decimal:
    if not s:
        return Decimal("0")
    s = str(s).strip()
    if not s:
        return Decimal("0")
    s = s.replace(",", ".")
    try:
        return Decimal(s)
    except Exception:
        return Decimal("0")


def _find_customfield(xml_text: str, name: str) -> str | None:
    m = re.search(
        rf'<CustomField\s+Name="{re.escape(name)}"\s+Value="([^"]*)"\s*/?>',
        xml_text or "",
        flags=re.IGNORECASE,
    )
    return m.group(1).strip() if m else None


def _find_customfieldrow_valor_por_desc(xml_text: str, contains_text: str) -> str | None:
    pat = (
        r'<CustomFieldRow\b[^>]*>.*?'
        r'Name="Descripcion"\s+Value="([^"]*)"\s*/>.*?'
        r'Name="Valor"\s+Value="([^"]+)"\s*/>.*?'
        r"</CustomFieldRow>"
    )
    for m in re.finditer(pat, xml_text or "", flags=re.IGNORECASE | re.DOTALL):
        desc = (m.group(1) or "").lower()
        val = (m.group(2) or "").strip()
        if contains_text.lower() in desc:
            return val
    return None


def _sumar_iva_porcentaje(root: ET.Element, ns: dict, pct_obj: float) -> float:
    total = 0.0

    subs_doc = root.findall("./cac:TaxTotal/cac:TaxSubtotal", ns)
    if subs_doc:
        for tax in subs_doc:
            amt = convertir_a_numero(obtener_texto(tax, "./cbc:TaxAmount", ns))
            pct_text = obtener_texto(tax, "./cac:TaxCategory/cbc:Percent", ns)
            try:
                pct = float(str(pct_text).strip())
                if abs(pct - pct_obj) < 0.01:
                    total += amt
            except Exception:
                continue
        return total

    subs_any = root.findall(".//cac:TaxTotal/cac:TaxSubtotal", ns)
    for tax in subs_any:
        amt = convertir_a_numero(obtener_texto(tax, "./cbc:TaxAmount", ns))
        pct_text = obtener_texto(tax, "./cac:TaxCategory/cbc:Percent", ns)
        try:
            pct = float(str(pct_text).strip())
            if abs(pct - pct_obj) < 0.01:
                total += amt
        except Exception:
            continue

    return total


def leer_datos_xml(path: str) -> dict | None:
    xml_text_for_regex = ""

    try:
        inner_xml = _extract_inner_invoice(path)
        if inner_xml:
            xml_text_for_regex = inner_xml
            try:
                root = ET.fromstring(inner_xml)
            except ET.ParseError:
                root = ET.fromstring(_clean_xml_text(inner_xml))
        else:
            with open(path, "rb") as f:
                raw = f.read()
            try:
                xml_text_for_regex = raw.decode("utf-8-sig", errors="replace")
            except Exception:
                xml_text_for_regex = raw.decode(errors="replace")

            xml_text_for_regex = _clean_xml_text(xml_text_for_regex)
            root = ET.fromstring(xml_text_for_regex)

            local = _local_name(root.tag)
            if local == "AttachedDocument":
                errores.append(f"AttachedDocument sin documento embebido: {os.path.basename(path)}")
                return None

    except ET.ParseError as e:
        errores.append(f"XML mal formado '{path}': {e}")
        return None
    except Exception as e:
        errores.append(f"Error al leer XML '{path}': {e}")
        return None

    ns = {
        "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
        "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    }

    emisor = obtener_texto(root, ".//cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name", ns)
    cliente = obtener_texto(root, ".//cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name", ns)

    if not cliente or cliente.lower() == "no aplica":
        cliente = obtener_texto(root, ".//cac:AccountingCustomerParty//cac:PartyLegalEntity/cbc:RegistrationName", ns)
    if not cliente:
        cliente = obtener_texto(root, ".//cac:AccountingCustomerParty//cac:PartyIdentification/cbc:ID", ns)

    numero = obtener_texto(root, "./cbc:ID", ns)

    descs = _extraer_descripciones_por_iva(root)
    descripcion_lineas = descs["all"]
    descripcion_iva19 = descs["iva19"]
    descripcion_iva5 = descs["iva5"]
    descripcion_iva0 = descs["iva0"]

    nit = obtener_texto(root, ".//cac:AccountingSupplierParty//cac:PartyLegalEntity/cbc:CompanyID", ns)
    tipo_contribuyente = obtener_texto(root, ".//cac:PartyTaxScheme/cbc:TaxLevelCode", ns)
    fecha_text = obtener_texto(root, "./cbc:IssueDate", ns)
    cufe = obtener_texto(root, ".//cbc:UUID", ns)

    ciudad_nombre = obtener_texto(
        root, ".//cac:AccountingSupplierParty//cac:PhysicalLocation//cac:Address//cbc:CityName", ns
    )
    ciudad_codigo = obtener_texto(
        root, ".//cac:AccountingSupplierParty//cac:PhysicalLocation//cac:Address//cbc:ID", ns
    )

    subtotal = convertir_a_numero(obtener_texto(root, ".//cac:LegalMonetaryTotal/cbc:LineExtensionAmount", ns))
    total_base = convertir_a_numero(obtener_texto(root, ".//cac:LegalMonetaryTotal/cbc:PayableAmount", ns))

    act_eco_el = root.find(".//{*}IndustryClassificationCode")
    actividad_economica = (act_eco_el.text.strip() if act_eco_el is not None and act_eco_el.text else "")

    if not actividad_economica:
        raw_xml = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8", errors="ignore")
        m = re.search(r"(?:CIIU|Actividad\s+Econ[oó]mica)[^\d]*(\d{4,5})", raw_xml, re.IGNORECASE)
        if m:
            actividad_economica = m.group(1)
    if not actividad_economica:
        actividad_economica = _extraer_actividad_de_pdf(path)

    iva_5 = _sumar_iva_porcentaje(root, ns, 5.0)
    iva_19 = _sumar_iva_porcentaje(root, ns, 19.0)

    reteiva = reteica = rete_fuente = 0.0
    for tax in root.findall("./cac:WithholdingTaxTotal/cac:TaxSubtotal", ns):
        amt = convertir_a_numero(obtener_texto(tax, "./cbc:TaxAmount", ns))
        tax_id = obtener_texto(tax, "./cac:TaxCategory/cac:TaxScheme/cbc:ID", ns).strip().lower()
        tax_name = obtener_texto(tax, "./cac:TaxCategory/cac:TaxScheme/cbc:Name", ns).strip().lower()

        if tax_id == "05" or "iva" in tax_name:
            reteiva += amt
        elif tax_id == "06" or "fuente" in tax_name or "renta" in tax_name:
            rete_fuente += amt
        elif tax_id == "07" or "ica" in tax_name:
            reteica += amt

    reteiva = -abs(reteiva)
    reteica = -abs(reteica)
    rete_fuente = -abs(rete_fuente)

    total_calc = total_base + reteiva + reteica + rete_fuente

    try:
        ajuste_retefuente = _find_customfieldrow_valor_por_desc(xml_text_for_regex, "retefuente")
        ajuste_notas = _find_customfield(xml_text_for_regex, "Ajuste_Notas_Credito")
        ajuste = ajuste_retefuente or ajuste_notas

        valor_total_pagar = _find_customfield(xml_text_for_regex, "Valor_Total_Pagar")

        if ajuste:
            aj = float(_dec_from_str(ajuste))
            rete_fuente = aj
            total_calc = total_base + reteiva + reteica + rete_fuente

        if valor_total_pagar:
            total_calc = float(_dec_from_str(valor_total_pagar))

    except Exception as e:
        errores.append(f"Error aplicando CustomFields (ajustes) en {os.path.basename(path)}: {e}")

    return {
        "Archivo": os.path.basename(path),
        "Empresa emisora": emisor,
        "CUFE": cufe,
        "Ciudad emisora": ciudad_nombre,
        "Código ciudad": ciudad_codigo,
        "NIT": nit,
        "Cliente": cliente,
        "Número de factura": numero,
        "Año": (fecha_text or "")[:4],
        "Mes": (fecha_text or "")[5:7],
        "Día": (fecha_text or "")[8:10],
        "Tipo de contribuyente": tipo_contribuyente,
        "Actividad económica": actividad_economica,
        "DescripcionLineas": descripcion_lineas,
        "DescripcionIVA19": descripcion_iva19,
        "DescripcionIVA5": descripcion_iva5,
        "DescripcionIVA0": descripcion_iva0,
        "Subtotal": subtotal,
        "IVA 5%": iva_5,
        "IVA 19%": iva_19,
        "Retención de IVA": reteiva,
        "Retención de ICA": reteica,
        "Retención en la fuente": rete_fuente,
        "Total": total_calc,
    }


def procesar_xml_en_carpeta(ruta_carpeta: str) -> tuple[list[dict], int]:
    registros: list[dict] = []
    errores_zip = 0

    for archivo in os.listdir(ruta_carpeta):
        if archivo.lower().endswith(".xml"):
            full_path = os.path.join(ruta_carpeta, archivo)
            datos = leer_datos_xml(full_path)
            if datos:
                registros.append(datos)
                print(f"✅ Procesado: {archivo}")
            else:
                errores_zip += 1

    return registros, errores_zip
