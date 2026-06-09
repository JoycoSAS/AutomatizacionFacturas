# services/factura_service.py
import base64
import os
import re
import xml.etree.ElementTree as ET
from decimal import Decimal
from html import unescape
from typing import Dict, List, Optional, Tuple

try:
    import PyPDF2  # pip install PyPDF2
except Exception:
    PyPDF2 = None

from utils.helpers import obtener_texto, convertir_a_numero

try:
    from utils.helpers import extraer_inner_invoice as _helper_extraer_inner_invoice
except Exception:
    _helper_extraer_inner_invoice = None

try:
    from utils.helpers import obtener_actividad_economica as _helper_obtener_actividad_economica
except Exception:
    _helper_obtener_actividad_economica = None

try:
    from utils.pdf_utils import extraer_texto_pdf as _extraer_texto_pdf_pdfminer
except Exception:
    _extraer_texto_pdf_pdfminer = None

from utils.logger import errores


PDF_FALLBACK_ENABLED = True


# ============================================================
# XML helpers robustos
# ============================================================

_CTRL_REGEX = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_AMP_FIX = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+|#x[0-9A-Fa-f]+);)")
_XML_DOC_RE = re.compile(
    r"(<\s*(?:Invoice|CreditNote|DebitNote)\b[\s\S]*?</\s*(?:Invoice|CreditNote|DebitNote)\s*>)",
    flags=re.IGNORECASE,
)


def _to_str(valor) -> str:
    if valor is None:
        return ""
    return str(valor).replace("\xa0", " ").strip()


def _clean_xml_text(txt: str) -> str:
    txt = _CTRL_REGEX.sub("", txt or "")
    txt = txt.replace("\ufeff", "")
    txt = _AMP_FIX.sub("&amp;", txt)
    return txt.strip()


def _clean_text(txt: str) -> str:
    txt = unescape(_to_str(txt))
    txt = txt.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    txt = re.sub(r"\s+", " ", txt).strip()
    if txt.upper() in {"NAN", "NONE", "NULL", "N/A", "NA"}:
        return ""
    return txt


def _local_name(tag: str) -> str:
    tag = _to_str(tag)
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    if ":" in tag:
        return tag.rsplit(":", 1)[-1]
    return tag


def _safe_read_text(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()

    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw.decode(enc, errors="replace")
        except Exception:
            continue

    return raw.decode("utf-8", errors="ignore")


def _safe_parse_xml(path: str) -> ET.Element:
    text = _safe_read_text(path)

    try:
        return ET.fromstring(text)
    except ET.ParseError:
        text2 = _clean_xml_text(text)
        return ET.fromstring(text2)


def _parse_xml_text(xml_text: str) -> ET.Element:
    xml_text = _clean_xml_text(xml_text)

    try:
        return ET.fromstring(xml_text)
    except ET.ParseError:
        xml_text = unescape(xml_text)
        xml_text = _clean_xml_text(xml_text)
        return ET.fromstring(xml_text)


def _extraer_xml_por_regex(texto: str) -> Optional[str]:
    texto = unescape(_to_str(texto))
    texto = _clean_xml_text(texto)

    m = _XML_DOC_RE.search(texto)
    if not m:
        return None

    return _clean_xml_text(m.group(1))


def _parece_base64(texto: str) -> bool:
    texto = re.sub(r"\s+", "", _to_str(texto))
    if len(texto) < 80:
        return False
    if len(texto) % 4 != 0:
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9+/=]+", texto))


def _decode_base64_texto(texto: str) -> str:
    limpio = re.sub(r"\s+", "", _to_str(texto))
    if not limpio:
        return ""

    try:
        raw = base64.b64decode(limpio, validate=False)
    except Exception:
        return ""

    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw.decode(enc, errors="replace")
        except Exception:
            continue

    return raw.decode("utf-8", errors="ignore")


def _validar_inner_invoice(xml_text: str) -> Optional[str]:
    """
    Retorna XML limpio si corresponde a Invoice/CreditNote/DebitNote.
    """
    if not xml_text:
        return None

    xml_text = _clean_xml_text(unescape(xml_text))

    inner = _extraer_xml_por_regex(xml_text)
    if inner:
        xml_text = inner

    try:
        root = _parse_xml_text(xml_text)
        local = _local_name(root.tag).lower()
        if local in {"invoice", "creditnote", "debitnote"}:
            return xml_text
    except Exception:
        return None

    return None


def _extract_inner_invoice(path: str) -> Optional[str]:
    """
    Extrae factura real si el XML es AttachedDocument.

    Soporta:
    - EmbeddedDocumentBinaryObject base64.
    - ExternalReference/Description con XML en CDATA.
    - Description con XML escapado.
    - Description con XML en base64.
    - URI apuntando a otro XML en la misma carpeta.
    - Helper robusto de utils.helpers.extraer_inner_invoice().
    """
    # 1) Usar helper robusto si existe.
    if _helper_extraer_inner_invoice is not None:
        try:
            helper_inner = _helper_extraer_inner_invoice(path)
            valid = _validar_inner_invoice(helper_inner or "")
            if valid:
                return valid
        except Exception as e:
            errores.append(f"Helper extraer_inner_invoice falló en {os.path.basename(path)}: {e}")

    try:
        root = _safe_parse_xml(path)
        original_text = _safe_read_text(path)
    except Exception as e:
        errores.append(f"Error extrayendo XML interno en {os.path.basename(path)}: {e}")
        return None

    # Si el archivo ya es Invoice/CreditNote/DebitNote, retornarlo completo.
    if _local_name(root.tag).lower() in {"invoice", "creditnote", "debitnote"}:
        return _clean_xml_text(original_text)

    # 2) Buscar Invoice directo en todo el texto.
    valid = _validar_inner_invoice(original_text)
    if valid:
        return valid

    # 3) EmbeddedDocumentBinaryObject.
    for bin_el in root.findall(".//{*}EmbeddedDocumentBinaryObject"):
        if bin_el is not None and bin_el.text:
            decoded = _decode_base64_texto(bin_el.text)
            valid = _validar_inner_invoice(decoded)
            if valid:
                return valid

    # 4) Description.
    desc_nodes = []
    for ruta in (
        ".//{*}Attachment//{*}ExternalReference//{*}Description",
        ".//{*}ExternalReference//{*}Description",
        ".//{*}Description",
    ):
        try:
            desc_nodes.extend(root.findall(ruta))
        except Exception:
            pass

    for desc in desc_nodes:
        if desc is None or not desc.text:
            continue

        raw = desc.text.strip()

        valid = _validar_inner_invoice(raw)
        if valid:
            return valid

        raw_unescaped = unescape(raw)
        valid = _validar_inner_invoice(raw_unescaped)
        if valid:
            return valid

        if _parece_base64(raw):
            decoded = _decode_base64_texto(raw)
            valid = _validar_inner_invoice(decoded)
            if valid:
                return valid

    # 5) URI a XML local.
    uri = root.find(".//{*}Attachment//{*}ExternalReference//{*}URI")
    if uri is not None and uri.text:
        maybe = uri.text.strip()
        if maybe and not re.match(r"^[a-z]+://", maybe, flags=re.I):
            carpeta = os.path.dirname(path)
            destino = os.path.join(carpeta, os.path.basename(maybe))
            if os.path.exists(destino) and destino.lower().endswith(".xml"):
                try:
                    txt = _safe_read_text(destino)
                    valid = _validar_inner_invoice(txt)
                    if valid:
                        return valid
                except Exception as e:
                    errores.append(f"No se pudo abrir URI '{maybe}' en {os.path.basename(path)}: {e}")

    return None


def _find_first_text(elem: ET.Element, paths: List[str], ns: Optional[dict] = None) -> str:
    for path in paths:
        try:
            txt = obtener_texto(elem, path, ns or {})
            txt = _clean_text(txt)
            if txt:
                return txt
        except Exception:
            pass

        try:
            nodo = elem.find(path, ns or {})
            if nodo is not None and nodo.text:
                txt = _clean_text(nodo.text)
                if txt:
                    return txt
        except Exception:
            pass

    return ""


def _find_all_texts_by_local(elem: ET.Element, local_name: str) -> List[str]:
    out = []
    if elem is None:
        return out

    for nodo in elem.iter():
        if _local_name(nodo.tag).lower() == local_name.lower() and nodo.text:
            txt = _clean_text(nodo.text)
            if txt:
                out.append(txt)

    return out


def _dedup_join(items: List[str]) -> str:
    seen = set()
    out = []

    for item in items:
        item = _clean_text(item)
        if not item:
            continue

        key = item.upper()
        if key in seen:
            continue

        seen.add(key)
        out.append(item)

    return "; ".join(out)


# ============================================================
# Líneas, descripciones y D1 por IVA
# ============================================================

def _iter_lineas_documento(root: ET.Element) -> List[ET.Element]:
    for xp in (".//{*}InvoiceLine", ".//{*}CreditNoteLine", ".//{*}DebitNoteLine"):
        lines = root.findall(xp)
        if lines:
            return lines
    return []


def _linea_descripcion(linea: ET.Element) -> str:
    """
    Extrae descripción robusta por línea.

    Mejora casos donde solo viene:
    - Item/Description
    - Item/Name
    - Note
    - SellersItemIdentification/ID
    - StandardItemIdentification/ID
    """
    candidatos: List[str] = []

    for ruta in (
        ".//{*}Item/{*}Description",
        ".//{*}Item/{*}Name",
        ".//{*}Description",
        ".//{*}Name",
        ".//{*}Note",
        ".//{*}SellersItemIdentification/{*}ID",
        ".//{*}StandardItemIdentification/{*}ID",
        ".//{*}BuyersItemIdentification/{*}ID",
    ):
        try:
            for nodo in linea.findall(ruta):
                if nodo is not None and nodo.text:
                    txt = _clean_text(nodo.text)
                    if txt:
                        candidatos.append(txt)
        except Exception:
            continue

    # Evitar que IDs numéricos puros se vuelvan la descripción si hay algo mejor.
    no_numericos = [x for x in candidatos if not re.fullmatch(r"\d{1,30}", x)]
    if no_numericos:
        return _dedup_join(no_numericos[:4])

    return _dedup_join(candidatos[:2])


def _linea_iva_percent(linea: ET.Element) -> Optional[float]:
    rutas = (
        ".//{*}TaxTotal/{*}TaxSubtotal/{*}TaxCategory/{*}Percent",
        ".//{*}ClassifiedTaxCategory/{*}Percent",
        ".//{*}TaxCategory/{*}Percent",
    )

    for ruta in rutas:
        try:
            pct = linea.find(ruta)
            if pct is not None and pct.text:
                return float(str(pct.text).strip().replace(",", "."))
        except Exception:
            continue

    return None


def _extraer_descripciones_por_iva(root: ET.Element) -> dict:
    """
    Retorna:
    - all: todas las descripciones.
    - iva19: productos con IVA 19%.
    - iva5: productos con IVA 5%.
    - iva0: productos sin IVA o sin porcentaje.
    """
    all_items: List[str] = []
    iva19: List[str] = []
    iva5: List[str] = []
    iva0: List[str] = []

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


# ============================================================
# Datos de cabecera XML
# ============================================================

def _party_supplier(root: ET.Element) -> Optional[ET.Element]:
    return root.find(".//{*}AccountingSupplierParty//{*}Party")


def _party_customer(root: ET.Element) -> Optional[ET.Element]:
    return root.find(".//{*}AccountingCustomerParty//{*}Party")


def _extraer_nombre_party(party: Optional[ET.Element]) -> str:
    if party is None:
        return ""

    paths = [
        ".//{*}PartyName/{*}Name",
        ".//{*}PartyLegalEntity/{*}RegistrationName",
        ".//{*}PartyTaxScheme/{*}RegistrationName",
        ".//{*}RegistrationName",
        ".//{*}Name",
    ]

    for path in paths:
        try:
            for nodo in party.findall(path):
                txt = _clean_text(nodo.text if nodo is not None else "")
                if txt:
                    return txt
        except Exception:
            pass

    return ""


def _extraer_nit_party(party: Optional[ET.Element]) -> str:
    if party is None:
        return ""

    paths = [
        ".//{*}PartyTaxScheme/{*}CompanyID",
        ".//{*}PartyLegalEntity/{*}CompanyID",
        ".//{*}PartyIdentification/{*}ID",
        ".//{*}CompanyID",
    ]

    for path in paths:
        try:
            for nodo in party.findall(path):
                txt = _clean_text(nodo.text if nodo is not None else "")
                if txt:
                    return re.sub(r"[^\d]", "", txt)
        except Exception:
            pass

    return ""


def _extraer_tax_level_party(party: Optional[ET.Element]) -> str:
    if party is None:
        return ""

    vals = []
    for path in (
        ".//{*}PartyTaxScheme/{*}TaxLevelCode",
        ".//{*}TaxLevelCode",
    ):
        try:
            for nodo in party.findall(path):
                txt = _clean_text(nodo.text if nodo is not None else "")
                if txt:
                    vals.append(txt)
        except Exception:
            pass

    return ";".join(dict.fromkeys(vals))


def _extraer_ciudad_party(party: Optional[ET.Element]) -> Tuple[str, str]:
    if party is None:
        return "", ""

    ciudad = ""
    codigo = ""

    # Prioridad: PhysicalLocation/Address.
    for base in (
        ".//{*}PhysicalLocation//{*}Address",
        ".//{*}RegistrationAddress",
        ".//{*}Address",
    ):
        try:
            addr = party.find(base)
            if addr is None:
                continue

            ciudad = _find_first_text(addr, [
                ".//{*}CityName",
                ".//{*}CitySubdivisionName",
            ])

            codigo = _find_first_text(addr, [
                ".//{*}ID",
            ])

            if ciudad or codigo:
                break
        except Exception:
            continue

    return ciudad, codigo


def _extraer_uuid(root: ET.Element) -> str:
    vals = _find_all_texts_by_local(root, "UUID")

    for v in vals:
        c = re.sub(r"[^0-9a-fA-F]", "", v).lower()
        if len(c) >= 64:
            return c[:96] if len(c) >= 96 else c

    return vals[0] if vals else ""


def _extraer_fecha(root: ET.Element, ns: dict) -> str:
    return _find_first_text(root, [
        "./cbc:IssueDate",
        "./{*}IssueDate",
        ".//{*}IssueDate",
    ], ns)


def _extraer_numero(root: ET.Element, ns: dict) -> str:
    numero = _find_first_text(root, [
        "./cbc:ID",
        "./{*}ID",
    ], ns)

    return numero


def _extraer_actividad(root: ET.Element, path: str) -> str:
    actividad = ""

    if _helper_obtener_actividad_economica is not None:
        try:
            actividad = _helper_obtener_actividad_economica(root)
        except Exception:
            actividad = ""

    if not actividad:
        for ruta in (
            ".//{*}AccountingSupplierParty//{*}Party//{*}IndustryClassificationCode",
            ".//{*}IndustryClassificationCode",
        ):
            try:
                nodo = root.find(ruta)
                if nodo is not None and nodo.text:
                    actividad = _clean_text(nodo.text)
                    if actividad:
                        break
            except Exception:
                pass

    if not actividad:
        raw_xml = ET.tostring(root, encoding="utf-8", method="xml").decode("utf-8", errors="ignore")
        m = re.search(r"(?:CIIU|Actividad\s+Econ[oó]mica)[^\d]*(\d{4,5})", raw_xml, re.IGNORECASE)
        if m:
            actividad = m.group(1)

    if not actividad:
        actividad = _extraer_actividad_de_pdf(path)

    return actividad


# ============================================================
# Actividad desde PDF asociado
# ============================================================

def _extraer_actividad_de_pdf(xml_path: str) -> str:
    if not PDF_FALLBACK_ENABLED:
        return ""

    carpeta = os.path.dirname(xml_path)

    try:
        archivos = os.listdir(carpeta)
    except Exception:
        return ""

    for fn in archivos:
        if not fn.lower().endswith(".pdf"):
            continue

        pdf_path = os.path.join(carpeta, fn)
        text = ""

        # Primero pdfminer si está disponible.
        if _extraer_texto_pdf_pdfminer is not None:
            try:
                text = _extraer_texto_pdf_pdfminer(pdf_path) or ""
            except Exception:
                text = ""

        # Fallback PyPDF2.
        if not text and PyPDF2 is not None:
            try:
                reader = PyPDF2.PdfReader(pdf_path)
                parts = []
                for page in reader.pages:
                    parts.append(page.extract_text() or "")
                text = "\n".join(parts)
            except Exception as e:
                errores.append(f"Error leyendo PDF '{fn}': {e}")

        if not text:
            continue

        m = re.search(r"(?:CIIU|Actividad\s+Econ[oó]mica)[^\d]*(\d{4,5})", text, re.IGNORECASE)
        if m:
            return m.group(1)

    return ""


# ============================================================
# Totales, impuestos y retenciones
# ============================================================

def _dec_from_str(s: Optional[str]) -> Decimal:
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


def _find_customfield(xml_text: str, name: str) -> Optional[str]:
    m = re.search(
        rf'<CustomField\s+Name="{re.escape(name)}"\s+Value="([^"]*)"\s*/?>',
        xml_text or "",
        flags=re.IGNORECASE,
    )
    return m.group(1).strip() if m else None


def _find_customfieldrow_valor_por_desc(xml_text: str, contains_text: str) -> Optional[str]:
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
    """
    Suma IVA por porcentaje.

    Prioridad:
    1. TaxSubtotal a nivel documento.
    2. TaxSubtotal en líneas si no hay totales de documento.
    """
    total = 0.0

    def _sumar_subtotales(subs) -> float:
        acc = 0.0
        for tax in subs:
            amt = convertir_a_numero(obtener_texto(tax, ["./cbc:TaxAmount", "./{*}TaxAmount"], ns))

            pct_text = obtener_texto(tax, [
                "./cac:TaxCategory/cbc:Percent",
                "./{*}TaxCategory/{*}Percent",
                ".//{*}Percent",
            ], ns)

            try:
                pct = float(str(pct_text).strip().replace(",", "."))
            except Exception:
                continue

            if abs(pct - pct_obj) < 0.01:
                acc += amt

        return acc

    subs_doc = root.findall("./cac:TaxTotal/cac:TaxSubtotal", ns)
    if not subs_doc:
        subs_doc = root.findall("./{*}TaxTotal/{*}TaxSubtotal")

    if subs_doc:
        total = _sumar_subtotales(subs_doc)
        return float(total or 0.0)

    subs_any = root.findall(".//cac:TaxTotal/cac:TaxSubtotal", ns)
    if not subs_any:
        subs_any = root.findall(".//{*}TaxTotal/{*}TaxSubtotal")

    total = _sumar_subtotales(subs_any)
    return float(total or 0.0)


def _extraer_retenciones(root: ET.Element, ns: dict) -> Tuple[float, float, float]:
    """
    Retorna:
    - reteiva
    - reteica
    - retefuente

    Por compatibilidad con el Excel actual se devuelven negativas.
    """
    reteiva = 0.0
    reteica = 0.0
    rete_fuente = 0.0

    subtotales = root.findall("./cac:WithholdingTaxTotal/cac:TaxSubtotal", ns)
    if not subtotales:
        subtotales = root.findall("./{*}WithholdingTaxTotal/{*}TaxSubtotal")
    if not subtotales:
        subtotales = root.findall(".//{*}WithholdingTaxTotal/{*}TaxSubtotal")

    for tax in subtotales:
        amt = convertir_a_numero(obtener_texto(tax, ["./cbc:TaxAmount", "./{*}TaxAmount"], ns))

        tax_id = obtener_texto(tax, [
            "./cac:TaxCategory/cac:TaxScheme/cbc:ID",
            "./{*}TaxCategory/{*}TaxScheme/{*}ID",
            ".//{*}TaxScheme/{*}ID",
        ], ns).strip().lower()

        tax_name = obtener_texto(tax, [
            "./cac:TaxCategory/cac:TaxScheme/cbc:Name",
            "./{*}TaxCategory/{*}TaxScheme/{*}Name",
            ".//{*}TaxScheme/{*}Name",
        ], ns).strip().lower()

        # DIAN frecuente:
        # 05 = ReteIVA, 06 = ReteFuente/Renta, 07 = ReteICA.
        if tax_id == "05" or "iva" in tax_name:
            reteiva += amt
        elif tax_id == "07" or "ica" in tax_name:
            reteica += amt
        elif tax_id == "06" or "fuente" in tax_name or "renta" in tax_name:
            rete_fuente += amt

    return -abs(reteiva), -abs(reteica), -abs(rete_fuente)


def _extraer_totales(root: ET.Element, ns: dict, xml_text_for_regex: str, path: str) -> Dict[str, float]:
    subtotal = convertir_a_numero(obtener_texto(root, [
        ".//cac:LegalMonetaryTotal/cbc:LineExtensionAmount",
        ".//{*}LegalMonetaryTotal/{*}LineExtensionAmount",
        ".//{*}LineExtensionAmount",
    ], ns))

    payable_amount = convertir_a_numero(obtener_texto(root, [
        ".//cac:LegalMonetaryTotal/cbc:PayableAmount",
        ".//{*}LegalMonetaryTotal/{*}PayableAmount",
        ".//{*}PayableAmount",
    ], ns))

    tax_inclusive = convertir_a_numero(obtener_texto(root, [
        ".//cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount",
        ".//{*}LegalMonetaryTotal/{*}TaxInclusiveAmount",
        ".//{*}TaxInclusiveAmount",
    ], ns))

    iva_5 = _sumar_iva_porcentaje(root, ns, 5.0)
    iva_19 = _sumar_iva_porcentaje(root, ns, 19.0)

    reteiva, reteica, rete_fuente = _extraer_retenciones(root, ns)

    # Regla importante:
    # PayableAmount normalmente ya es el total final a pagar.
    # Solo calculamos si PayableAmount no viene.
    if payable_amount > 0:
        total_calc = payable_amount
    elif tax_inclusive > 0:
        total_calc = tax_inclusive + reteiva + reteica + rete_fuente
    else:
        total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

    # CustomFields de algunos proveedores.
    try:
        ajuste_retefuente = _find_customfieldrow_valor_por_desc(xml_text_for_regex, "retefuente")
        ajuste_notas = _find_customfield(xml_text_for_regex, "Ajuste_Notas_Credito")
        ajuste = ajuste_retefuente or ajuste_notas

        valor_total_pagar = _find_customfield(xml_text_for_regex, "Valor_Total_Pagar")

        if ajuste:
            aj = float(_dec_from_str(ajuste))
            rete_fuente = aj

            # Si no hay PayableAmount explícito, recalcular; si sí existe, se conserva.
            if payable_amount <= 0:
                total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

        if valor_total_pagar:
            total_calc = float(_dec_from_str(valor_total_pagar))

    except Exception as e:
        errores.append(f"Error aplicando CustomFields (ajustes) en {os.path.basename(path)}: {e}")

    return {
        "Subtotal": float(subtotal or 0.0),
        "IVA 5%": float(iva_5 or 0.0),
        "IVA 19%": float(iva_19 or 0.0),
        "Retención de IVA": float(reteiva or 0.0),
        "Retención de ICA": float(reteica or 0.0),
        "Retención en la fuente": float(rete_fuente or 0.0),
        "Total": float(total_calc or 0.0),
    }


# ============================================================
# Lectura principal
# ============================================================


# ============================================================
# Refuerzos 2026-05-08 para XML UBL / AttachedDocument
# ============================================================

def _score_party_20260508(party: Optional[ET.Element]) -> int:
    """
    Puntúa un Party para elegir el más completo cuando el XML trae
    varias estructuras o viene dentro de AttachedDocument.
    """
    if party is None:
        return -999

    score = 0

    nombre = _extraer_nombre_party(party)
    nit = _extraer_nit_party(party)
    ciudad, codigo = _extraer_ciudad_party(party)
    tax_level = _extraer_tax_level_party(party)

    if nombre:
        score += 50
    if nit:
        score += 35
    if ciudad:
        score += 20
    if codigo:
        score += 10
    if tax_level:
        score += 10

    flat = _clean_text(" ".join([nombre, nit, ciudad, codigo, tax_level]))
    if re.search(r"(?i)\b(tel[eé]fono|contacto|correo|email)\b", flat):
        score -= 30

    return score


def _best_party_20260508(root: ET.Element, party_container_local: str) -> Optional[ET.Element]:
    """
    Selecciona el Party más completo dentro de AccountingSupplierParty
    o AccountingCustomerParty. Esto evita tomar nodos incompletos
    cuando el XML viene como AttachedDocument o trae anexos.
    """
    candidates: List[ET.Element] = []

    for node in root.iter():
        if _local_name(node.tag).lower() == party_container_local.lower():
            for child in node.iter():
                if _local_name(child.tag).lower() == "party":
                    candidates.append(child)

    if not candidates:
        return None

    return max(candidates, key=_score_party_20260508)


def _party_supplier_20260508(root: ET.Element) -> Optional[ET.Element]:
    return _best_party_20260508(root, "AccountingSupplierParty")


def _party_customer_20260508(root: ET.Element) -> Optional[ET.Element]:
    return _best_party_20260508(root, "AccountingCustomerParty")


def _find_invoice_root_20260508(root: ET.Element) -> ET.Element:
    """
    Retorna el Invoice/CreditNote/DebitNote real.
    Normalmente _extract_inner_invoice ya devuelve el XML interno, pero
    este refuerzo protege casos donde llegue un contenedor.
    """
    local = _local_name(root.tag).lower()

    if local in {"invoice", "creditnote", "debitnote"}:
        return root

    for node in root.iter():
        if _local_name(node.tag).lower() in {"invoice", "creditnote", "debitnote"}:
            return node

    return root


def _extraer_uuid_invoice_20260508(root: ET.Element) -> str:
    """
    Toma UUID/CUFE preferiblemente del Invoice real.
    Evita CUFE falso generado desde textos mezclados.
    """
    invoice = _find_invoice_root_20260508(root)

    for path in ("./{*}UUID", ".//{*}UUID"):
        try:
            for nodo in invoice.findall(path):
                txt = _clean_text(nodo.text if nodo is not None else "")
                c = re.sub(r"[^0-9a-fA-F]", "", txt).lower()

                if len(c) >= 96:
                    return c[:96]
                if len(c) >= 64:
                    return c
        except Exception:
            pass

    return ""


def _extraer_numero_invoice_20260508(root: ET.Element, ns: dict) -> str:
    invoice = _find_invoice_root_20260508(root)
    return _find_first_text(invoice, [
        "./cbc:ID",
        "./{*}ID",
    ], ns)


def _extraer_fecha_invoice_20260508(root: ET.Element, ns: dict) -> str:
    invoice = _find_invoice_root_20260508(root)
    return _find_first_text(invoice, [
        "./cbc:IssueDate",
        "./{*}IssueDate",
        ".//{*}IssueDate",
    ], ns)


def _extraer_ciudad_party_mejorada_20260508(party: Optional[ET.Element]) -> Tuple[str, str]:
    """
    Extrae ciudad y código de ciudad evitando capturar textos tipo:
    'TELÉFONO', 'CONTACTO', 'EMAIL'.
    """
    ciudad, codigo = _extraer_ciudad_party(party)

    if ciudad and re.search(r"(?i)\b(tel[eé]fono|contacto|email|correo)\b", ciudad):
        ciudad = ""

    if codigo and not re.fullmatch(r"\d{4,8}", codigo):
        codigo = ""

    if ciudad or codigo:
        return ciudad, codigo

    if party is None:
        return "", ""

    for base in (
        ".//{*}PhysicalLocation//{*}Address",
        ".//{*}RegistrationAddress",
        ".//{*}PostalAddress",
        ".//{*}Address",
    ):
        try:
            for addr in party.findall(base):
                city = _find_first_text(addr, [
                    "./{*}CityName",
                    ".//{*}CityName",
                ])
                code = _find_first_text(addr, [
                    "./{*}ID",
                    ".//{*}ID",
                ])

                if city and re.search(r"(?i)\b(tel[eé]fono|contacto|email|correo)\b", city):
                    city = ""

                if code and not re.fullmatch(r"\d{4,8}", code):
                    code = ""

                if city or code:
                    return city, code
        except Exception:
            pass

    return "", ""


def _extraer_actividad_mejorada_20260508(root: ET.Element, supplier: Optional[ET.Element], path: str) -> str:
    """
    Actividad económica:
    1. Busca en el supplier real.
    2. Usa la lógica existente.
    3. Usa fallback PDF asociado.
    """
    if supplier is not None:
        for ruta in (
            ".//{*}IndustryClassificationCode",
            ".//{*}CorporateRegistrationScheme/{*}ID",
        ):
            try:
                for nodo in supplier.findall(ruta):
                    txt = _clean_text(nodo.text if nodo is not None else "")
                    if re.fullmatch(r"\d{4,5}", txt):
                        return txt
            except Exception:
                pass

    actividad = _extraer_actividad(root, path)

    if actividad:
        return actividad

    return ""


def _sumar_impuesto_porcentaje_20260508(
    root: ET.Element,
    ns: dict,
    pct_obj: float,
    incluir_no_iva: bool = False,
) -> float:
    """
    Suma impuestos por porcentaje.
    Por defecto solo toma IVA para evitar meter INC 8% como IVA 19%.
    """
    total = 0.0

    def tax_scheme_is_iva(taxsubtotal: ET.Element) -> bool:
        tax_id = obtener_texto(taxsubtotal, [
            ".//cac:TaxScheme/cbc:ID",
            ".//{*}TaxScheme/{*}ID",
        ], ns).strip().lower()

        tax_name = obtener_texto(taxsubtotal, [
            ".//cac:TaxScheme/cbc:Name",
            ".//{*}TaxScheme/{*}Name",
        ], ns).strip().lower()

        if incluir_no_iva:
            return True

        # DIAN normalmente usa 01 = IVA. INC suele ser 04.
        if tax_id == "01":
            return True

        if "iva" in tax_name:
            return True

        return False

    subtotales: List[ET.Element] = []

    for xp in (
        "./cac:TaxTotal/cac:TaxSubtotal",
        "./{*}TaxTotal/{*}TaxSubtotal",
        ".//cac:TaxTotal/cac:TaxSubtotal",
        ".//{*}TaxTotal/{*}TaxSubtotal",
    ):
        try:
            if "cac:" in xp:
                subtotales.extend(root.findall(xp, ns))
            else:
                subtotales.extend(root.findall(xp))
        except Exception:
            pass

    vistos = set()
    unicos: List[ET.Element] = []

    for st in subtotales:
        marker = id(st)
        if marker in vistos:
            continue
        vistos.add(marker)
        unicos.append(st)

    for tax in unicos:
        pct_text = obtener_texto(tax, [
            "./cac:TaxCategory/cbc:Percent",
            "./{*}TaxCategory/{*}Percent",
            ".//{*}Percent",
        ], ns)

        try:
            pct = float(str(pct_text).strip().replace(",", "."))
        except Exception:
            continue

        if abs(pct - pct_obj) >= 0.01:
            continue

        if not tax_scheme_is_iva(tax):
            continue

        amt = convertir_a_numero(obtener_texto(tax, [
            "./cbc:TaxAmount",
            "./{*}TaxAmount",
        ], ns))

        total += amt

    return float(total or 0.0)


def _extraer_totales_mejorado_20260508(
    root: ET.Element,
    ns: dict,
    xml_text_for_regex: str,
    path: str,
) -> Dict[str, float]:
    """
    Totales XML más seguros:
    - Usa LegalMonetaryTotal del Invoice real.
    - Usa PayableAmount como total final si existe.
    - No convierte INC 8% en IVA 19%.
    """
    invoice = _find_invoice_root_20260508(root)

    subtotal = convertir_a_numero(obtener_texto(invoice, [
        ".//cac:LegalMonetaryTotal/cbc:LineExtensionAmount",
        ".//{*}LegalMonetaryTotal/{*}LineExtensionAmount",
    ], ns))

    payable_amount = convertir_a_numero(obtener_texto(invoice, [
        ".//cac:LegalMonetaryTotal/cbc:PayableAmount",
        ".//{*}LegalMonetaryTotal/{*}PayableAmount",
    ], ns))

    tax_inclusive = convertir_a_numero(obtener_texto(invoice, [
        ".//cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount",
        ".//{*}LegalMonetaryTotal/{*}TaxInclusiveAmount",
    ], ns))

    iva_5 = _sumar_impuesto_porcentaje_20260508(invoice, ns, 5.0)
    iva_19 = _sumar_impuesto_porcentaje_20260508(invoice, ns, 19.0)

    reteiva, reteica, rete_fuente = _extraer_retenciones(invoice, ns)

    if payable_amount > 0:
        total_calc = payable_amount
    elif tax_inclusive > 0:
        total_calc = tax_inclusive + reteiva + reteica + rete_fuente
    else:
        total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

    try:
        ajuste_retefuente = _find_customfieldrow_valor_por_desc(xml_text_for_regex, "retefuente")
        ajuste_notas = _find_customfield(xml_text_for_regex, "Ajuste_Notas_Credito")
        ajuste = ajuste_retefuente or ajuste_notas

        valor_total_pagar = _find_customfield(xml_text_for_regex, "Valor_Total_Pagar")

        if ajuste:
            aj = float(_dec_from_str(ajuste))
            rete_fuente = aj

            if payable_amount <= 0:
                total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

        if valor_total_pagar:
            total_calc = float(_dec_from_str(valor_total_pagar))

    except Exception as e:
        errores.append(f"Error aplicando CustomFields (ajustes) en {os.path.basename(path)}: {e}")

    return {
        "Subtotal": float(subtotal or 0.0),
        "IVA 5%": float(iva_5 or 0.0),
        "IVA 19%": float(iva_19 or 0.0),
        "Retención de IVA": float(reteiva or 0.0),
        "Retención de ICA": float(reteica or 0.0),
        "Retención en la fuente": float(rete_fuente or 0.0),
        "Total": float(total_calc or 0.0),
    }


def leer_datos_xml(path: str) -> Optional[dict]:
    xml_text_for_regex = ""

    try:
        inner_xml = _extract_inner_invoice(path)

        if inner_xml:
            xml_text_for_regex = inner_xml
            root = _parse_xml_text(inner_xml)
        else:
            xml_text_for_regex = _safe_read_text(path)
            xml_text_for_regex = _clean_xml_text(xml_text_for_regex)
            root = _parse_xml_text(xml_text_for_regex)

            local = _local_name(root.tag)
            if local == "AttachedDocument":
                errores.append(f"AttachedDocument sin documento embebido: {os.path.basename(path)}")
                return None

        root = _find_invoice_root_20260508(root)

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

    supplier = _party_supplier_20260508(root) or _party_supplier(root)
    customer = _party_customer_20260508(root) or _party_customer(root)

    emisor = _extraer_nombre_party(supplier)
    cliente = _extraer_nombre_party(customer)

    if not cliente or cliente.lower() == "no aplica":
        cliente = _find_first_text(root, [
            ".//cac:AccountingCustomerParty//cac:PartyLegalEntity/cbc:RegistrationName",
            ".//{*}AccountingCustomerParty//{*}PartyLegalEntity/{*}RegistrationName",
            ".//cac:AccountingCustomerParty//cac:PartyName/cbc:Name",
            ".//{*}AccountingCustomerParty//{*}PartyName/{*}Name",
        ], ns)

    if not cliente:
        cliente = _extraer_nit_party(customer)

    numero = _extraer_numero_invoice_20260508(root, ns)

    descs = _extraer_descripciones_por_iva(root)
    descripcion_lineas = descs["all"]
    descripcion_iva19 = descs["iva19"]
    descripcion_iva5 = descs["iva5"]
    descripcion_iva0 = descs["iva0"]

    nit = _extraer_nit_party(supplier)
    tipo_contribuyente = _extraer_tax_level_party(supplier)

    fecha_text = _extraer_fecha_invoice_20260508(root, ns)
    cufe = _extraer_uuid_invoice_20260508(root)

    ciudad_nombre, ciudad_codigo = _extraer_ciudad_party_mejorada_20260508(supplier)

    actividad_economica = _extraer_actividad_mejorada_20260508(root, supplier, path)

    totales = _extraer_totales_mejorado_20260508(root, ns, xml_text_for_regex, path)

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
        "Subtotal": totales["Subtotal"],
        "IVA 5%": totales["IVA 5%"],
        "IVA 19%": totales["IVA 19%"],
        "Retención de IVA": totales["Retención de IVA"],
        "Retención de ICA": totales["Retención de ICA"],
        "Retención en la fuente": totales["Retención en la fuente"],
        "Total": totales["Total"],
    }


def procesar_xml_en_carpeta(ruta_carpeta: str) -> tuple[list[dict], int]:
    registros: list[dict] = []
    errores_zip = 0

    try:
        archivos = os.listdir(ruta_carpeta)
    except Exception as e:
        errores.append(f"No se pudo listar carpeta XML '{ruta_carpeta}': {e}")
        return registros, 1

    for archivo in archivos:
        if archivo.lower().endswith(".xml"):
            full_path = os.path.join(ruta_carpeta, archivo)
            datos = leer_datos_xml(full_path)

            if datos:
                registros.append(datos)
                print(f"✅ Procesado: {archivo}")
            else:
                errores_zip += 1

    return registros, errores_zip



# ============================================================
# PATCH 2026-05-11 - Refuerzo final XML / AttachedDocument
# ============================================================
# Se deja al final para no romper lo que ya funciona.
# En Python, esta definición final de leer_datos_xml reemplaza la anterior.
#
# Objetivo:
# - Leer Invoice/CreditNote/DebitNote real aunque venga embebido en AttachedDocument.
# - Evitar tomar fechas de vigencia/anexo como fecha de factura.
# - Evitar CUFE falso.
# - Tomar supplier/customer directos del documento real.
# - Evitar que INC 8% se cargue como IVA 19%.
# - Mantener descripciones por IVA para D1 y proveedores con varias líneas.
# ============================================================

_leer_datos_xml_pre_20260511 = leer_datos_xml


def _read_xml_text_20260511(path: str) -> str:
    try:
        return _safe_read_text(path)
    except Exception:
        with open(path, "rb") as fh:
            raw = fh.read()
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return raw.decode(enc, errors="replace")
            except Exception:
                pass
        return raw.decode("utf-8", errors="ignore")


def _extract_invoice_xml_document_20260511(path: str) -> Tuple[Optional[str], str]:
    """
    Devuelve:
    - XML del documento tributario real si lo encuentra.
    - Texto original leído.
    """
    original_text = _read_xml_text_20260511(path)

    # 1) Primero usar la extracción existente, que ya cubre base64/helper/CDATA.
    try:
        inner = _extract_inner_invoice(path)
        if inner:
            inner = _clean_xml_text(unescape(inner))
            m = _XML_DOC_RE.search(inner)
            if m:
                inner = _clean_xml_text(m.group(1))
            root_test = _parse_xml_text(inner)
            if _local_name(root_test.tag).lower() in {"invoice", "creditnote", "debitnote"}:
                return inner, original_text
    except Exception as e:
        errores.append(f"Extracción XML 20260511 falló usando helper en {os.path.basename(path)}: {e}")

    # 2) Buscar directamente el bloque <Invoice>, <CreditNote> o <DebitNote>
    #    dentro del texto completo. Esto corrige AttachedDocument con XML escapado.
    try:
        raw = _clean_xml_text(unescape(original_text))
        m = _XML_DOC_RE.search(raw)
        if m:
            inner = _clean_xml_text(m.group(1))
            root_test = _parse_xml_text(inner)
            if _local_name(root_test.tag).lower() in {"invoice", "creditnote", "debitnote"}:
                return inner, original_text
    except Exception:
        pass

    # 3) Si el archivo ya es factura real, usarlo completo.
    try:
        raw = _clean_xml_text(unescape(original_text))
        root_test = _parse_xml_text(raw)
        if _local_name(root_test.tag).lower() in {"invoice", "creditnote", "debitnote"}:
            return raw, original_text
    except Exception:
        pass

    return None, original_text


def _party_container_direct_20260511(root: ET.Element, container_local_name: str) -> Optional[ET.Element]:
    """
    Prioriza el contenedor directo del Invoice real.
    Evita capturar Parties de anexos o nodos secundarios.
    """
    for child in list(root):
        if _local_name(child.tag).lower() == container_local_name.lower():
            return child

    for node in root.iter():
        if _local_name(node.tag).lower() == container_local_name.lower():
            return node

    return None


def _party_direct_20260511(root: ET.Element, container_local_name: str) -> Optional[ET.Element]:
    cont = _party_container_direct_20260511(root, container_local_name)
    if cont is None:
        return None

    for child in list(cont):
        if _local_name(child.tag).lower() == "party":
            return child

    for node in cont.iter():
        if _local_name(node.tag).lower() == "party":
            return node

    return None


def _text_first_child_20260511(elem: Optional[ET.Element], local_names: List[str]) -> str:
    if elem is None:
        return ""

    wanted = {x.lower() for x in local_names}

    # Primero hijos directos.
    for child in list(elem):
        if _local_name(child.tag).lower() in wanted and child.text:
            txt = _clean_text(child.text)
            if txt:
                return txt

    # Luego descendientes.
    for node in elem.iter():
        if _local_name(node.tag).lower() in wanted and node.text:
            txt = _clean_text(node.text)
            if txt:
                return txt

    return ""


def _first_direct_text_20260511(root: ET.Element, local_name: str) -> str:
    """
    Toma campos directos de Invoice, no de anexos:
    ID, IssueDate, UUID.
    """
    lname = local_name.lower()

    for child in list(root):
        if _local_name(child.tag).lower() == lname and child.text:
            txt = _clean_text(child.text)
            if txt:
                return txt

    return ""


def _extraer_uuid_directo_20260511(root: ET.Element) -> str:
    cufe = _first_direct_text_20260511(root, "UUID")

    if not cufe:
        # Como fallback, primer UUID dentro del documento real.
        cufe = _text_first_child_20260511(root, ["UUID"])

    c = re.sub(r"[^0-9a-fA-F]", "", cufe or "").lower()
    if len(c) >= 96:
        return c[:96]
    if len(c) >= 64:
        return c

    # Si no cumple longitud de CUFE, no inventar CUFE.
    return ""


def _extraer_numero_directo_20260511(root: ET.Element) -> str:
    numero = _first_direct_text_20260511(root, "ID")
    return _clean_text(numero)


def _extraer_fecha_directa_20260511(root: ET.Element) -> str:
    """
    IssueDate directo del Invoice real.
    No toma vigencias ni fechas de anexos.
    """
    fecha = _first_direct_text_20260511(root, "IssueDate")
    return _clean_text(fecha)


def _extraer_nombre_party_20260511(party: Optional[ET.Element]) -> str:
    if party is None:
        return ""

    # Orden más seguro en DIAN: RegistrationName suele ser razón social real.
    for ruta in (
        ".//{*}PartyLegalEntity/{*}RegistrationName",
        ".//{*}PartyTaxScheme/{*}RegistrationName",
        ".//{*}PartyName/{*}Name",
        ".//{*}RegistrationName",
        ".//{*}Name",
    ):
        try:
            for node in party.findall(ruta):
                if node is not None and node.text:
                    txt = _clean_text(node.text)
                    if txt:
                        return txt
        except Exception:
            pass

    return ""


def _extraer_nit_party_20260511(party: Optional[ET.Element]) -> str:
    if party is None:
        return ""

    # Priorizar CompanyID sobre PartyIdentification/ID.
    for ruta in (
        ".//{*}PartyTaxScheme/{*}CompanyID",
        ".//{*}PartyLegalEntity/{*}CompanyID",
        ".//{*}CompanyID",
        ".//{*}PartyIdentification/{*}ID",
    ):
        try:
            for node in party.findall(ruta):
                if node is not None and node.text:
                    nit = re.sub(r"[^\d]", "", _clean_text(node.text))
                    if nit:
                        return nit
        except Exception:
            pass

    return ""


def _extraer_ciudad_party_20260511(party: Optional[ET.Element]) -> Tuple[str, str]:
    if party is None:
        return "", ""

    # Orden normal DIAN.
    for base in (
        ".//{*}PhysicalLocation//{*}Address",
        ".//{*}RegistrationAddress",
        ".//{*}PostalAddress",
        ".//{*}Address",
    ):
        try:
            for addr in party.findall(base):
                ciudad = _find_first_text(addr, [
                    "./{*}CityName",
                    ".//{*}CityName",
                ])
                codigo = _find_first_text(addr, [
                    "./{*}ID",
                    ".//{*}ID",
                ])

                if ciudad and re.search(r"(?i)\b(tel[eé]fono|contacto|correo|email)\b", ciudad):
                    ciudad = ""

                if codigo and not re.fullmatch(r"\d{4,8}", codigo):
                    codigo = ""

                if ciudad or codigo:
                    return ciudad, codigo
        except Exception:
            pass

    return "", ""


def _extraer_tax_level_party_20260511(party: Optional[ET.Element]) -> str:
    if party is None:
        return ""

    vals: List[str] = []

    for ruta in (
        ".//{*}PartyTaxScheme/{*}TaxLevelCode",
        ".//{*}TaxLevelCode",
    ):
        try:
            for node in party.findall(ruta):
                if node is not None and node.text:
                    txt = _clean_text(node.text)
                    if txt:
                        vals.append(txt)
        except Exception:
            pass

    return ";".join(dict.fromkeys(vals))


def _extraer_actividad_20260511(root: ET.Element, supplier: Optional[ET.Element], path: str) -> str:
    # Actividad no define calidad crítica, pero si viene, se toma.
    for base in (supplier, root):
        if base is None:
            continue
        for ruta in (
            ".//{*}IndustryClassificationCode",
            ".//{*}CorporateRegistrationScheme/{*}ID",
        ):
            try:
                for node in base.findall(ruta):
                    if node is not None and node.text:
                        txt = _clean_text(node.text)
                        if re.fullmatch(r"\d{4,5}", txt):
                            return txt
            except Exception:
                pass

    try:
        return _extraer_actividad(root, path) or ""
    except Exception:
        return ""


def _tax_scheme_is_iva_20260511(taxsubtotal: ET.Element, ns: dict) -> bool:
    tax_id = obtener_texto(taxsubtotal, [
        ".//cac:TaxScheme/cbc:ID",
        ".//{*}TaxScheme/{*}ID",
    ], ns).strip().lower()

    tax_name = obtener_texto(taxsubtotal, [
        ".//cac:TaxScheme/cbc:Name",
        ".//{*}TaxScheme/{*}Name",
    ], ns).strip().lower()

    # DIAN: 01 = IVA. INC suele ser 04.
    return tax_id == "01" or "iva" in tax_name


def _sumar_iva_porcentaje_20260511(root: ET.Element, ns: dict, pct_obj: float) -> float:
    """
    Suma IVA por porcentaje.
    No suma INC 8% ni otros impuestos diferentes de IVA.
    """
    total = 0.0
    vistos = set()

    # Priorizar TaxTotal directo del documento.
    subtotales: List[ET.Element] = []
    for tax_total in list(root):
        if _local_name(tax_total.tag).lower() != "taxtotal":
            continue
        for child in list(tax_total):
            if _local_name(child.tag).lower() == "taxsubtotal":
                subtotales.append(child)

    # Si no hay subtotales directos, buscar global.
    if not subtotales:
        for xp in (
            ".//cac:TaxTotal/cac:TaxSubtotal",
            ".//{*}TaxTotal/{*}TaxSubtotal",
        ):
            try:
                if "cac:" in xp:
                    subtotales.extend(root.findall(xp, ns))
                else:
                    subtotales.extend(root.findall(xp))
            except Exception:
                pass

    for tax in subtotales:
        marker = id(tax)
        if marker in vistos:
            continue
        vistos.add(marker)

        if not _tax_scheme_is_iva_20260511(tax, ns):
            continue

        pct_text = obtener_texto(tax, [
            "./cac:TaxCategory/cbc:Percent",
            "./{*}TaxCategory/{*}Percent",
            ".//{*}Percent",
        ], ns)

        try:
            pct = float(str(pct_text).strip().replace(",", "."))
        except Exception:
            continue

        if abs(pct - pct_obj) >= 0.01:
            continue

        amt = convertir_a_numero(obtener_texto(tax, [
            "./cbc:TaxAmount",
            "./{*}TaxAmount",
        ], ns))

        total += amt

    return float(total or 0.0)


def _extraer_totales_xml_20260511(root: ET.Element, ns: dict, xml_text_for_regex: str, path: str) -> Dict[str, float]:
    subtotal = convertir_a_numero(obtener_texto(root, [
        ".//cac:LegalMonetaryTotal/cbc:LineExtensionAmount",
        ".//{*}LegalMonetaryTotal/{*}LineExtensionAmount",
    ], ns))

    payable_amount = convertir_a_numero(obtener_texto(root, [
        ".//cac:LegalMonetaryTotal/cbc:PayableAmount",
        ".//{*}LegalMonetaryTotal/{*}PayableAmount",
    ], ns))

    tax_inclusive = convertir_a_numero(obtener_texto(root, [
        ".//cac:LegalMonetaryTotal/cbc:TaxInclusiveAmount",
        ".//{*}LegalMonetaryTotal/{*}TaxInclusiveAmount",
    ], ns))

    iva_5 = _sumar_iva_porcentaje_20260511(root, ns, 5.0)
    iva_19 = _sumar_iva_porcentaje_20260511(root, ns, 19.0)

    reteiva, reteica, rete_fuente = _extraer_retenciones(root, ns)

    # PayableAmount normalmente ya es el total final.
    if payable_amount > 0:
        total_calc = payable_amount
    elif tax_inclusive > 0:
        total_calc = tax_inclusive + reteiva + reteica + rete_fuente
    else:
        total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

    try:
        ajuste_retefuente = _find_customfieldrow_valor_por_desc(xml_text_for_regex, "retefuente")
        ajuste_notas = _find_customfield(xml_text_for_regex, "Ajuste_Notas_Credito")
        ajuste = ajuste_retefuente or ajuste_notas

        valor_total_pagar = _find_customfield(xml_text_for_regex, "Valor_Total_Pagar")

        if ajuste:
            aj = float(_dec_from_str(ajuste))
            rete_fuente = aj
            if payable_amount <= 0:
                total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

        if valor_total_pagar:
            total_calc = float(_dec_from_str(valor_total_pagar))

    except Exception as e:
        errores.append(f"Error aplicando CustomFields 20260511 en {os.path.basename(path)}: {e}")

    return {
        "Subtotal": float(subtotal or 0.0),
        "IVA 5%": float(iva_5 or 0.0),
        "IVA 19%": float(iva_19 or 0.0),
        "Retención de IVA": float(reteiva or 0.0),
        "Retención de ICA": float(reteica or 0.0),
        "Retención en la fuente": float(rete_fuente or 0.0),
        "Total": float(total_calc or 0.0),
    }


def _descripciones_por_iva_20260511(root: ET.Element) -> Dict[str, str]:
    """
    Usa la lógica existente, pero agrega limpieza básica para evitar descripciones vacías
    cuando el XML tiene nombres en Item/Name o Description.
    """
    try:
        descs = _extraer_descripciones_por_iva(root) or {}
    except Exception:
        descs = {}

    if descs.get("all"):
        return {
            "all": descs.get("all", ""),
            "iva19": descs.get("iva19", ""),
            "iva5": descs.get("iva5", ""),
            "iva0": descs.get("iva0", ""),
        }

    all_items: List[str] = []
    iva19: List[str] = []
    iva5: List[str] = []
    iva0: List[str] = []

    for linea in _iter_lineas_documento(root):
        desc = _linea_descripcion(linea)
        desc = _clean_text(desc)

        if not desc:
            continue

        if desc not in all_items:
            all_items.append(desc)

        pct = _linea_iva_percent(linea)

        if pct is None or abs(pct) < 0.0001:
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


def leer_datos_xml(path: str) -> Optional[dict]:
    xml_text_for_regex = ""

    try:
        xml_doc, original_text = _extract_invoice_xml_document_20260511(path)

        if not xml_doc:
            # Último fallback: comportamiento anterior.
            return _leer_datos_xml_pre_20260511(path)

        xml_text_for_regex = xml_doc
        root = _parse_xml_text(xml_doc)
        root = _find_invoice_root_20260508(root)

        if _local_name(root.tag).lower() not in {"invoice", "creditnote", "debitnote"}:
            errores.append(f"XML sin documento tributario real 20260511: {os.path.basename(path)}")
            return None

    except ET.ParseError as e:
        errores.append(f"XML mal formado 20260511 '{path}': {e}")
        return None
    except Exception as e:
        errores.append(f"Error al leer XML 20260511 '{path}': {e}")
        return None

    ns = {
        "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
        "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    }

    supplier = _party_direct_20260511(root, "AccountingSupplierParty") or _party_supplier_20260508(root) or _party_supplier(root)
    customer = _party_direct_20260511(root, "AccountingCustomerParty") or _party_customer_20260508(root) or _party_customer(root)

    emisor = _extraer_nombre_party_20260511(supplier)
    cliente = _extraer_nombre_party_20260511(customer)

    if not cliente:
        cliente = _extraer_nit_party_20260511(customer)

    numero = _extraer_numero_directo_20260511(root)
    fecha_text = _extraer_fecha_directa_20260511(root)
    cufe = _extraer_uuid_directo_20260511(root)

    ciudad_nombre, ciudad_codigo = _extraer_ciudad_party_20260511(supplier)

    nit = _extraer_nit_party_20260511(supplier)
    tipo_contribuyente = _extraer_tax_level_party_20260511(supplier)
    actividad_economica = _extraer_actividad_20260511(root, supplier, path)

    descs = _descripciones_por_iva_20260511(root)
    totales = _extraer_totales_xml_20260511(root, ns, xml_text_for_regex, path)

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
        "DescripcionLineas": descs["all"],
        "DescripcionIVA19": descs["iva19"],
        "DescripcionIVA5": descs["iva5"],
        "DescripcionIVA0": descs["iva0"],
        "Subtotal": totales["Subtotal"],
        "IVA 5%": totales["IVA 5%"],
        "IVA 19%": totales["IVA 19%"],
        "Retención de IVA": totales["Retención de IVA"],
        "Retención de ICA": totales["Retención de ICA"],
        "Retención en la fuente": totales["Retención en la fuente"],
        "Total": totales["Total"],
    }


# =====================================================================
# PATCH 2026-05-13 - Refuerzo XML lote PARCIAL/MINIMA
# =====================================================================
# Objetivo de este bloque:
# - Mejorar facturas que vienen en ZIP/XML, sin tocar el flujo de correo.
# - Corregir descripciones en XML DIAN donde se estaban mezclando textos
#   como "IVA", notas o códigos internos.
# - Corregir lectura de valores monetarios con punto decimal XML:
#     4621.85 -> 4621.85, NO 4.62
# - Normalizar NIT sin dígito de verificación cuando venga con guion:
#     890.903.407-9 -> 890903407
# - Mantener CUFE real únicamente si viene como UUID válido.
#
# Casos cubiertos del lote:
# - TDG TECHNOLOGIES SAS / FE25442: descripción ELECTRODOS 45X45MM; ENVIO.
# - CRYSTAL S.A.S / CM04140574: descripción desde líneas XML.
# - FERRETERIA CATATUMBO / FE123102: cabecera, descripción y valores reales.
# =====================================================================

_leer_datos_xml_pre_20260513 = leer_datos_xml


def _nit_sin_dv_20260513(value: str) -> str:
    """
    Normaliza NIT/CC quitando separadores y dígito de verificación si viene
    explícitamente después de guion.

    Ejemplos:
    - 890.903.407-9 -> 890903407
    - 1.098.738.506-0 -> 1098738506
    - 901416941 -> 901416941
    """
    s = _to_str(value)
    if not s:
        return ""

    m = re.search(r"([0-9][0-9\.\s]{5,})\s*-\s*\d\b", s)
    if m:
        return re.sub(r"[^0-9]", "", m.group(1))

    return re.sub(r"[^0-9]", "", s)


def _money_xml_20260513(value) -> float:
    """
    Conversor robusto para importes XML.

    Reglas:
    - XML DIAN suele venir con punto decimal: 4621.85.
    - Excel/PDF colombiano puede venir 4.621,85.
    - No divide miles por error.
    """
    if value is None:
        return 0.0

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return float(value)
        except Exception:
            return 0.0

    s = _clean_text(value)
    if not s:
        return 0.0

    neg = False
    if "(" in s and ")" in s:
        neg = True
    if re.search(r"^\s*-", s):
        neg = True

    s = s.replace("\xa0", " ")
    s = re.sub(r"(?i)\b(COP|USD|EUR|PESOS?|DOLARES?|DÓLARES?)\b", "", s)
    s = s.replace("$", "").replace("(", "").replace(")", "")
    s = re.sub(r"[^0-9,\.\-]", "", s)
    s = s.replace("-", "")

    if not s:
        return 0.0

    if "," in s and "." in s:
        # El separador decimal es el último que aparece.
        if s.rfind(",") > s.rfind("."):
            # 4.621,85
            s = s.replace(".", "").replace(",", ".")
        else:
            # 4,621.85
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) > 2:
            if len(parts[-1]) in {1, 2}:
                s = "".join(parts[:-1]) + "." + parts[-1]
            else:
                s = "".join(parts)
        else:
            if len(parts[-1]) in {1, 2}:
                s = parts[0].replace(".", "") + "." + parts[-1]
            else:
                s = "".join(parts)
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2:
            if len(parts[-1]) in {1, 2}:
                s = "".join(parts[:-1]) + "." + parts[-1]
            else:
                s = "".join(parts)
        else:
            # Caso ambiguo:
            # - 92.900 normalmente es miles => 92900
            # - 4621.85 es decimal XML => 4621.85
            # - 100239.49 es decimal XML => 100239.49
            entero, dec = parts[0], parts[1]
            if len(dec) == 3 and len(entero) <= 3:
                s = entero + dec
            else:
                s = entero + "." + dec

    try:
        val = float(s)
        return -val if neg and val > 0 else val
    except Exception:
        return 0.0


def _find_text_20260513(elem: Optional[ET.Element], paths: List[str], ns: Optional[dict] = None) -> str:
    if elem is None:
        return ""

    ns = ns or {}

    for path in paths:
        try:
            node = elem.find(path, ns)
            if node is not None and node.text:
                txt = _clean_text(node.text)
                if txt:
                    return txt
        except Exception:
            pass

    return ""


def _find_money_20260513(elem: Optional[ET.Element], paths: List[str], ns: Optional[dict] = None) -> float:
    txt = _find_text_20260513(elem, paths, ns)
    return _money_xml_20260513(txt)


def _extraer_nit_party_20260513(party: Optional[ET.Element]) -> str:
    if party is None:
        return ""

    for ruta in (
        ".//{*}PartyTaxScheme/{*}CompanyID",
        ".//{*}PartyLegalEntity/{*}CompanyID",
        ".//{*}CompanyID",
        ".//{*}PartyIdentification/{*}ID",
    ):
        try:
            for node in party.findall(ruta):
                if node is not None and node.text:
                    nit = _nit_sin_dv_20260513(node.text)
                    if nit:
                        return nit
        except Exception:
            pass

    return ""


def _limpiar_descripcion_xml_20260513(value: str) -> str:
    txt = _clean_text(value)
    if not txt:
        return ""

    txt = re.sub(r"\s+", " ", txt).strip(" ;|-")

    upper = txt.upper().strip()

    # Ruido frecuente al usar búsquedas globales dentro de InvoiceLine.
    if upper in {
        "IVA", "INC", "NO APLICA", "N/A", "NA", "NULL", "NONE", "NAN",
        "01", "04", "05", "06", "07", "94", "NIU", "UND", "UNIDAD",
    }:
        return ""

    # Evitar IDs numéricos puros como descripción.
    if re.fullmatch(r"\d{1,30}", upper):
        return ""

    return txt


def _texto_item_identificacion_20260513(item: Optional[ET.Element]) -> str:
    if item is None:
        return ""

    candidatos: List[str] = []
    for ruta in (
        ".//{*}SellersItemIdentification/{*}ID",
        ".//{*}StandardItemIdentification/{*}ID",
        ".//{*}BuyersItemIdentification/{*}ID",
    ):
        try:
            for node in item.findall(ruta):
                if node is not None and node.text:
                    val = _limpiar_descripcion_xml_20260513(node.text)
                    if not val:
                        continue
                    up = val.upper()
                    # Se aceptan códigos descriptivos tipo LLAMAROJA.
                    # Se descartan E1, 66102, códigos de barras, etc.
                    if len(up) < 4:
                        continue
                    if re.fullmatch(r"\d+", up):
                        continue
                    if re.fullmatch(r"[A-Z]\d+", up):
                        continue
                    candidatos.append(val)
        except Exception:
            pass

    return candidatos[0] if candidatos else ""


def _linea_descripcion_20260513(linea: ET.Element) -> str:
    """
    Extrae descripción limpia por línea.

    Importante:
    - Busca principalmente dentro de cac:Item.
    - No usa .//Name global porque eso captura TaxScheme/Name = IVA.
    - Combina código descriptivo + descripción cuando aplica:
      LLAMAROJA + DESTAPADOR... -> LLAMAROJA DESTAPADOR...
    """
    item = None
    try:
        item = linea.find(".//{*}Item")
    except Exception:
        item = None

    candidatos: List[str] = []

    if item is not None:
        for ruta in (
            "./{*}Description",
            "./{*}Name",
            ".//{*}Description",
            ".//{*}Name",
        ):
            try:
                for node in item.findall(ruta):
                    if node is not None and node.text:
                        val = _limpiar_descripcion_xml_20260513(node.text)
                        if val:
                            candidatos.append(val)
            except Exception:
                pass

    # Fallback: usar Note solo si no hay Item/Description.
    if not candidatos:
        try:
            for node in linea.findall("./{*}Note"):
                if node is not None and node.text:
                    val = _limpiar_descripcion_xml_20260513(node.text)
                    if val:
                        candidatos.append(val)
        except Exception:
            pass

    desc = _dedup_join(candidatos)
    codigo = _texto_item_identificacion_20260513(item)

    if codigo and desc:
        if not desc.upper().startswith(codigo.upper()):
            return _clean_text(f"{codigo} {desc}")
        return desc

    return desc or codigo


def _descripciones_por_iva_20260513(root: ET.Element) -> Dict[str, str]:
    all_items: List[str] = []
    iva19: List[str] = []
    iva5: List[str] = []
    iva0: List[str] = []

    for linea in _iter_lineas_documento(root):
        desc = _linea_descripcion_20260513(linea)
        desc = _limpiar_descripcion_xml_20260513(desc)
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


def _tax_scheme_is_iva_20260513(taxsubtotal: ET.Element, ns: dict) -> bool:
    tax_id = _find_text_20260513(taxsubtotal, [
        ".//cac:TaxScheme/cbc:ID",
        ".//{*}TaxScheme/{*}ID",
    ], ns).strip().lower()

    tax_name = _find_text_20260513(taxsubtotal, [
        ".//cac:TaxScheme/cbc:Name",
        ".//{*}TaxScheme/{*}Name",
    ], ns).strip().lower()

    return tax_id == "01" or "iva" in tax_name


def _sumar_iva_porcentaje_20260513(root: ET.Element, ns: dict, pct_obj: float) -> float:
    total = 0.0
    vistos = set()
    subtotales: List[ET.Element] = []

    # Priorizar TaxTotal directo del documento para evitar duplicar impuestos de líneas.
    for tax_total in list(root):
        if _local_name(tax_total.tag).lower() != "taxtotal":
            continue
        for child in list(tax_total):
            if _local_name(child.tag).lower() == "taxsubtotal":
                subtotales.append(child)

    if not subtotales:
        for xp in (
            ".//cac:TaxTotal/cac:TaxSubtotal",
            ".//{*}TaxTotal/{*}TaxSubtotal",
        ):
            try:
                if "cac:" in xp:
                    subtotales.extend(root.findall(xp, ns))
                else:
                    subtotales.extend(root.findall(xp))
            except Exception:
                pass

    for tax in subtotales:
        marker = id(tax)
        if marker in vistos:
            continue
        vistos.add(marker)

        if not _tax_scheme_is_iva_20260513(tax, ns):
            continue

        pct_text = _find_text_20260513(tax, [
            "./cac:TaxCategory/cbc:Percent",
            "./{*}TaxCategory/{*}Percent",
            ".//{*}Percent",
        ], ns)

        try:
            pct = float(str(pct_text).strip().replace(",", "."))
        except Exception:
            continue

        if abs(pct - pct_obj) >= 0.01:
            continue

        total += _find_money_20260513(tax, [
            "./cbc:TaxAmount",
            "./{*}TaxAmount",
        ], ns)

    return float(total or 0.0)


def _extraer_retenciones_20260513(root: ET.Element, ns: dict) -> Tuple[float, float, float]:
    reteiva = 0.0
    reteica = 0.0
    rete_fuente = 0.0

    subtotales: List[ET.Element] = []

    for xp in (
        "./cac:WithholdingTaxTotal/cac:TaxSubtotal",
        "./{*}WithholdingTaxTotal/{*}TaxSubtotal",
        ".//{*}WithholdingTaxTotal/{*}TaxSubtotal",
    ):
        try:
            if "cac:" in xp:
                subtotales.extend(root.findall(xp, ns))
            else:
                subtotales.extend(root.findall(xp))
        except Exception:
            pass

    vistos = set()
    for tax in subtotales:
        marker = id(tax)
        if marker in vistos:
            continue
        vistos.add(marker)

        amt = _find_money_20260513(tax, [
            "./cbc:TaxAmount",
            "./{*}TaxAmount",
        ], ns)

        tax_id = _find_text_20260513(tax, [
            "./cac:TaxCategory/cac:TaxScheme/cbc:ID",
            "./{*}TaxCategory/{*}TaxScheme/{*}ID",
            ".//{*}TaxScheme/{*}ID",
        ], ns).strip().lower()

        tax_name = _find_text_20260513(tax, [
            "./cac:TaxCategory/cac:TaxScheme/cbc:Name",
            "./{*}TaxCategory/{*}TaxScheme/{*}Name",
            ".//{*}TaxScheme/{*}Name",
        ], ns).strip().lower()

        if tax_id == "05" or "iva" in tax_name:
            reteiva += amt
        elif tax_id == "07" or "ica" in tax_name:
            reteica += amt
        elif tax_id == "06" or "fuente" in tax_name or "renta" in tax_name:
            rete_fuente += amt

    return -abs(reteiva), -abs(reteica), -abs(rete_fuente)


def _extraer_totales_xml_20260513(root: ET.Element, ns: dict, xml_text_for_regex: str, path: str) -> Dict[str, float]:
    monetary = None
    try:
        monetary = root.find(".//cac:LegalMonetaryTotal", ns)
        if monetary is None:
            monetary = root.find(".//{*}LegalMonetaryTotal")
    except Exception:
        monetary = None

    subtotal = _find_money_20260513(monetary, [
        "./cbc:LineExtensionAmount",
        "./{*}LineExtensionAmount",
    ], ns)

    payable_amount = _find_money_20260513(monetary, [
        "./cbc:PayableAmount",
        "./{*}PayableAmount",
    ], ns)

    tax_inclusive = _find_money_20260513(monetary, [
        "./cbc:TaxInclusiveAmount",
        "./{*}TaxInclusiveAmount",
    ], ns)

    iva_5 = _sumar_iva_porcentaje_20260513(root, ns, 5.0)
    iva_19 = _sumar_iva_porcentaje_20260513(root, ns, 19.0)
    reteiva, reteica, rete_fuente = _extraer_retenciones_20260513(root, ns)

    if payable_amount > 0:
        total_calc = payable_amount
    elif tax_inclusive > 0:
        total_calc = tax_inclusive + reteiva + reteica + rete_fuente
    else:
        total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

    try:
        ajuste_retefuente = _find_customfieldrow_valor_por_desc(xml_text_for_regex, "retefuente")
        ajuste_notas = _find_customfield(xml_text_for_regex, "Ajuste_Notas_Credito")
        ajuste = ajuste_retefuente or ajuste_notas

        valor_total_pagar = _find_customfield(xml_text_for_regex, "Valor_Total_Pagar")

        if ajuste:
            rete_fuente = _money_xml_20260513(ajuste)
            if payable_amount <= 0:
                total_calc = subtotal + iva_5 + iva_19 + reteiva + reteica + rete_fuente

        if valor_total_pagar:
            total_calc = _money_xml_20260513(valor_total_pagar)

    except Exception as e:
        errores.append(f"Error aplicando CustomFields 20260513 en {os.path.basename(path)}: {e}")

    return {
        "Subtotal": float(subtotal or 0.0),
        "IVA 5%": float(iva_5 or 0.0),
        "IVA 19%": float(iva_19 or 0.0),
        "Retención de IVA": float(reteiva or 0.0),
        "Retención de ICA": float(reteica or 0.0),
        "Retención en la fuente": float(rete_fuente or 0.0),
        "Total": float(total_calc or 0.0),
    }


def _extraer_numero_directo_20260513(root: ET.Element) -> str:
    numero = _extraer_numero_directo_20260511(root)
    return _clean_text(numero)


def leer_datos_xml(path: str) -> Optional[dict]:
    """
    Lector XML final activo 2026-05-13.
    Devuelve el mismo contrato de datos que el servicio original.
    """
    xml_text_for_regex = ""

    try:
        xml_doc, original_text = _extract_invoice_xml_document_20260511(path)

        if not xml_doc:
            return _leer_datos_xml_pre_20260513(path)

        xml_text_for_regex = xml_doc
        root = _parse_xml_text(xml_doc)
        root = _find_invoice_root_20260508(root)

        if _local_name(root.tag).lower() not in {"invoice", "creditnote", "debitnote"}:
            errores.append(f"XML sin documento tributario real 20260513: {os.path.basename(path)}")
            return None

    except ET.ParseError as e:
        errores.append(f"XML mal formado 20260513 '{path}': {e}")
        return None
    except Exception as e:
        errores.append(f"Error al leer XML 20260513 '{path}': {e}")
        return None

    ns = {
        "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
        "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
    }

    supplier = _party_direct_20260511(root, "AccountingSupplierParty") or _party_supplier_20260508(root) or _party_supplier(root)
    customer = _party_direct_20260511(root, "AccountingCustomerParty") or _party_customer_20260508(root) or _party_customer(root)

    emisor = _extraer_nombre_party_20260511(supplier)
    cliente = _extraer_nombre_party_20260511(customer)

    if not cliente:
        cliente = _extraer_nit_party_20260513(customer)

    numero = _extraer_numero_directo_20260513(root)
    fecha_text = _extraer_fecha_directa_20260511(root)
    cufe = _extraer_uuid_directo_20260511(root)

    ciudad_nombre, ciudad_codigo = _extraer_ciudad_party_20260511(supplier)

    nit = _extraer_nit_party_20260513(supplier)
    tipo_contribuyente = _extraer_tax_level_party_20260511(supplier)
    actividad_economica = _extraer_actividad_20260511(root, supplier, path)

    descs = _descripciones_por_iva_20260513(root)
    totales = _extraer_totales_xml_20260513(root, ns, xml_text_for_regex, path)

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
        "DescripcionLineas": descs.get("all", ""),
        "DescripcionIVA19": descs.get("iva19", ""),
        "DescripcionIVA5": descs.get("iva5", ""),
        "DescripcionIVA0": descs.get("iva0", ""),
        "Subtotal": totales["Subtotal"],
        "IVA 5%": totales["IVA 5%"],
        "IVA 19%": totales["IVA 19%"],
        "Retención de IVA": totales["Retención de IVA"],
        "Retención de ICA": totales["Retención de ICA"],
        "Retención en la fuente": totales["Retención en la fuente"],
        "Total": totales["Total"],
    }


# =====================================================================
# PATCH 2026-05-13-C - Procesamiento recursivo de XML en carpetas extraídas
# =====================================================================
# Motivo:
# Algunos ZIP de proveedores vienen con una subcarpeta interna, por ejemplo:
#   carpeta_zip/fv09014169410002600004051/FES-FE123102.xml
# La función original procesar_xml_en_carpeta() solo revisaba el primer nivel
# con os.listdir(ruta_carpeta), por eso devolvía 0 registros aunque el XML
# sí existiera. Esta definición final reemplaza la anterior y mantiene el
# mismo contrato: retorna (registros, errores_zip).
# =====================================================================


def _iter_xmls_recursivo_20260513(ruta_carpeta: str):
    """
    Itera XMLs dentro de ruta_carpeta de forma recursiva.

    - Ordena carpetas y archivos para resultados determinísticos.
    - Ignora carpetas ocultas y marcadores internos.
    - No procesa archivos que no sean .xml.
    """
    if not ruta_carpeta or not os.path.isdir(ruta_carpeta):
        return

    for root_dir, dirnames, filenames in os.walk(ruta_carpeta):
        dirnames[:] = sorted(
            d for d in dirnames
            if d and not d.startswith('.') and d.lower() not in {'__pycache__'}
        )

        for filename in sorted(filenames):
            if not filename.lower().endswith('.xml'):
                continue

            full_path = os.path.join(root_dir, filename)
            yield full_path


def procesar_xml_en_carpeta(ruta_carpeta: str) -> tuple[list[dict], int]:
    """
    Procesa todos los XML encontrados dentro de una carpeta extraída.

    Versión final activa 2026-05-13-C:
    - Busca XMLs de forma recursiva con os.walk().
    - Corrige ZIPs que traen una subcarpeta interna.
    - Mantiene compatibilidad con los ZIPs que ya venían planos.
    - Usa leer_datos_xml() final activo, por lo que conserva los parches
      de XML/AttachedDocument/descripciones/valores del 2026-05-13.
    """
    registros: list[dict] = []
    errores_zip = 0

    if not ruta_carpeta or not os.path.isdir(ruta_carpeta):
        errores.append(f"No se pudo listar carpeta XML '{ruta_carpeta}': no existe o no es carpeta")
        return registros, 1

    try:
        xml_paths = list(_iter_xmls_recursivo_20260513(ruta_carpeta) or [])
    except Exception as e:
        errores.append(f"No se pudo listar carpeta XML recursiva '{ruta_carpeta}': {e}")
        return registros, 1

    if not xml_paths:
        return registros, 0

    for full_path in xml_paths:
        rel_name = os.path.relpath(full_path, ruta_carpeta)

        try:
            datos = leer_datos_xml(full_path)
        except Exception as e:
            errores.append(f"Error procesando XML '{full_path}': {e}")
            datos = None

        if datos:
            # Mantener Archivo como nombre base del XML, igual que antes,
            # para que el Excel no quede con rutas largas.
            datos["Archivo"] = os.path.basename(full_path)
            registros.append(datos)
            print(f"✅ Procesado: {rel_name}")
        else:
            errores_zip += 1

    return registros, errores_zip

