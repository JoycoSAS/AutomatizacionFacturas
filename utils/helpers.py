import base64
import re
import xml.etree.ElementTree as ET
from html import unescape
from pathlib import Path
from typing import Iterable, Optional, Union


# ============================================================
# Helpers generales XML / números
# ============================================================

_XML_DOC_RE = re.compile(
    r"(<\s*(?:Invoice|CreditNote|DebitNote)\b[\s\S]*?</\s*(?:Invoice|CreditNote|DebitNote)\s*>)",
    flags=re.IGNORECASE,
)

_CTRL_REGEX = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_AMP_FIX = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+|#x[0-9A-Fa-f]+);)")


def _to_str(valor) -> str:
    if valor is None:
        return ""
    return str(valor).replace("\xa0", " ").strip()


def _limpiar_texto_xml(texto: str) -> str:
    """
    Limpieza segura para poder parsear XML que viene con caracteres de control
    o ampersands sueltos.
    """
    texto = _to_str(texto)
    if not texto:
        return ""

    texto = _CTRL_REGEX.sub("", texto)
    texto = texto.replace("\ufeff", "")
    texto = _AMP_FIX.sub("&amp;", texto)
    return texto.strip()


def _local_name(tag: str) -> str:
    """
    Convierte '{namespace}Invoice' o 'cbc:ID' en 'Invoice' / 'ID'.
    """
    tag = _to_str(tag)
    if not tag:
        return ""

    if "}" in tag:
        return tag.rsplit("}", 1)[-1]

    if ":" in tag:
        return tag.rsplit(":", 1)[-1]

    return tag


def _texto_limpio(valor) -> str:
    """
    Normaliza texto de nodos XML sin destruir espacios útiles.
    """
    s = _to_str(valor)
    if not s:
        return ""

    s = unescape(s)
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()

    if s.upper() in {"NAN", "NONE", "NULL"}:
        return ""

    return s


def convertir_a_numero(texto):
    """
    Convierte valores numéricos en distintos formatos a float.

    Soporta:
    - 398.000,00
    - 252,100.84
    - 300,000.00
    - $2.321.874
    - COP 1.120.630
    - -63.271,72
    - (63.271,72)
    - 0,00

    Si no puede convertir, retorna 0.0.
    """
    if texto is None:
        return 0.0

    if isinstance(texto, (int, float)):
        try:
            return float(texto)
        except Exception:
            return 0.0

    s = _to_str(texto)
    if not s:
        return 0.0

    s_upper = s.upper()
    if s_upper in {"NAN", "NONE", "NULL", "NO APLICA", "N/A"}:
        return 0.0

    negativo = False
    if "(" in s and ")" in s:
        negativo = True

    # Quitar símbolos y palabras frecuentes de moneda.
    s = s.replace("\xa0", " ")
    s = re.sub(r"(?i)\b(COP|USD|EUR|PESOS?|DOLARES?|DÓLARES?)\b", "", s)
    s = s.replace("$", "")
    s = s.replace("(", "").replace(")", "")
    s = s.strip()

    if s.startswith("-"):
        negativo = True

    # Mantener solo dígitos, separadores y signo.
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if not s or s in {"-", ".", ","}:
        return 0.0

    # Si hay varios signos menos, dejar solo el primero como indicador.
    if "-" in s:
        negativo = True
        s = s.replace("-", "")

    if not s:
        return 0.0

    # Caso con coma y punto: el separador decimal suele ser el último.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            # Formato latino: 1.234.567,89
            s = s.replace(".", "").replace(",", ".")
        else:
            # Formato anglo: 1,234,567.89
            s = s.replace(",", "")

    elif "," in s:
        partes = s.split(",")

        # Si la última parte tiene 1 o 2 dígitos, se asume decimal.
        # Si no, se asume separador de miles.
        if len(partes) > 1 and len(partes[-1]) in {1, 2}:
            s = "".join(partes[:-1]) + "." + partes[-1]
        else:
            s = "".join(partes)

    elif "." in s:
        partes = s.split(".")

        # Si hay varios puntos y el último no parece decimal, son miles.
        if len(partes) > 2:
            if len(partes[-1]) in {1, 2}:
                s = "".join(partes[:-1]) + "." + partes[-1]
            else:
                s = "".join(partes)
        else:
            # Un solo punto:
            # 300.00 => decimal
            # 300.000 => miles
            if len(partes[-1]) == 3 and len(partes[0]) <= 3:
                s = "".join(partes)

    try:
        val = float(s)
        return -val if negativo and val > 0 else val
    except Exception:
        return 0.0


def _iter_elementos_por_local_name(elem, local_name: str):
    """
    Itera nodos por nombre local, ignorando namespace/prefijo.
    """
    if elem is None:
        return

    objetivo = _to_str(local_name)
    if not objetivo:
        return

    objetivo = objetivo.lower()
    for nodo in elem.iter():
        if _local_name(nodo.tag).lower() == objetivo:
            yield nodo


def _find_text_by_path_or_local(elem, tag: str, ns=None, default: str = "") -> str:
    """
    Búsqueda compatible con ElementTree:
    1) intenta elem.find(tag, ns)
    2) intenta elem.find(tag)
    3) si es un nombre simple, busca por local-name ignorando namespace
    """
    if elem is None:
        return default

    tag = _to_str(tag)
    if not tag:
        return default

    candidatos = []

    try:
        nodo = elem.find(tag, ns or {})
        if nodo is not None:
            candidatos.append(nodo)
    except Exception:
        pass

    try:
        nodo = elem.find(tag)
        if nodo is not None:
            candidatos.append(nodo)
    except Exception:
        pass

    # Si viene como ruta con {namespace} o .//{*}, ElementTree ya intentó arriba.
    # Como fallback, buscar por el último segmento local.
    tag_limpio = tag.split("/")[-1]
    tag_limpio = tag_limpio.replace("{*}", "")
    tag_limpio = _local_name(tag_limpio)

    if tag_limpio:
        try:
            candidatos.extend(list(_iter_elementos_por_local_name(elem, tag_limpio)))
        except Exception:
            pass

    for nodo in candidatos:
        txt = _texto_limpio(getattr(nodo, "text", ""))
        if txt:
            return txt

    return default


def obtener_texto(elem, tag, ns=None, default=""):
    """
    Obtiene texto de un nodo XML de forma tolerante a namespaces.

    Mantiene compatibilidad con el uso anterior:
        obtener_texto(elem, tag, ns, default="")

    Además soporta:
        - tag como lista/tupla de rutas alternativas
        - búsqueda por local-name si el namespace no coincide
    """
    if isinstance(tag, (list, tuple, set)):
        for t in tag:
            txt = obtener_texto(elem, t, ns, default="")
            if txt:
                return txt
        return default

    return _find_text_by_path_or_local(elem, str(tag or ""), ns=ns, default=default)


def obtener_actividad_economica(root):
    """
    Extrae Actividad económica del emisor.

    En UBL normalmente viene en:
    AccountingSupplierParty / Party / IndustryClassificationCode

    Si no aparece ahí, intenta buscar cualquier IndustryClassificationCode.
    """
    if root is None:
        return ""

    rutas = [
        ".//{*}AccountingSupplierParty//{*}Party//{*}IndustryClassificationCode",
        ".//{*}AccountingSupplierParty//{*}IndustryClassificationCode",
        ".//{*}IndustryClassificationCode",
    ]

    for ruta in rutas:
        try:
            nodo = root.find(ruta)
            txt = _texto_limpio(nodo.text if nodo is not None else "")
            if txt:
                return txt
        except Exception:
            pass

    return ""


# ============================================================
# AttachedDocument / Invoice embebido
# ============================================================

def _leer_archivo_texto(path: Union[str, Path]) -> str:
    data = Path(path).read_bytes()

    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return data.decode(enc, errors="replace")
        except Exception:
            continue

    return data.decode("utf-8", errors="ignore")


def _extraer_xml_por_regex(texto: str) -> Optional[str]:
    texto = _to_str(texto)
    if not texto:
        return None

    texto = unescape(texto)
    texto = _limpiar_texto_xml(texto)

    m = _XML_DOC_RE.search(texto)
    if not m:
        return None

    inner = _limpiar_texto_xml(m.group(1))
    return inner if inner else None


def _parece_base64(s: str) -> bool:
    s = _to_str(s)
    if len(s) < 80:
        return False

    limpio = re.sub(r"\s+", "", s)
    if len(limpio) < 80:
        return False

    if len(limpio) % 4 != 0:
        return False

    return bool(re.fullmatch(r"[A-Za-z0-9+/=]+", limpio))


def _decode_base64_a_texto(s: str) -> str:
    limpio = re.sub(r"\s+", "", _to_str(s))
    if not limpio:
        return ""

    try:
        raw = base64.b64decode(limpio, validate=True)
    except Exception:
        return ""

    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw.decode(enc, errors="replace")
        except Exception:
            continue

    return raw.decode("utf-8", errors="ignore")


def _buscar_descripciones_xml(root) -> Iterable[str]:
    """
    Busca textos candidatos donde la DIAN suele guardar el XML real:
    AttachedDocument/Attachment/ExternalReference/Description.
    También revisa cualquier Description por si el namespace varía.
    """
    if root is None:
        return []

    candidatos = []

    rutas = [
        ".//{*}Attachment//{*}ExternalReference//{*}Description",
        ".//{*}ExternalReference//{*}Description",
        ".//{*}Description",
    ]

    for ruta in rutas:
        try:
            for nodo in root.findall(ruta):
                txt = _to_str(nodo.text)
                if txt:
                    candidatos.append(txt)
        except Exception:
            pass

    return candidatos


def _validar_xml_factura(xml_text: str) -> Optional[str]:
    """
    Retorna XML limpio si corresponde a Invoice/CreditNote/DebitNote parseable.
    """
    xml_text = _limpiar_texto_xml(xml_text)
    if not xml_text:
        return None

    inner = _extraer_xml_por_regex(xml_text)
    if inner:
        xml_text = inner

    try:
        root = ET.fromstring(xml_text)
        local = _local_name(root.tag).lower()
        if local in {"invoice", "creditnote", "debitnote"}:
            return xml_text
    except Exception:
        # Si no parsea, pero regex logró detectar estructura, igual puede
        # servir después de una limpieza adicional.
        inner = _extraer_xml_por_regex(xml_text)
        if inner:
            try:
                root = ET.fromstring(inner)
                local = _local_name(root.tag).lower()
                if local in {"invoice", "creditnote", "debitnote"}:
                    return inner
            except Exception:
                pass

    return None


def extraer_inner_invoice(path):
    """
    Extrae el XML interno si está incrustado dentro de AttachedDocument.

    Soporta:
    - CDATA con <Invoice>...</Invoice>
    - XML escapado con &lt;Invoice&gt;...&lt;/Invoice&gt;
    - Description con contenido base64
    - CreditNote / DebitNote además de Invoice

    Retorna:
        str con XML interno limpio, o None si no existe.
    """
    try:
        texto = _leer_archivo_texto(path)
    except Exception:
        return None

    # 1) Si el archivo directamente contiene un Invoice/CreditNote/DebitNote,
    # retornarlo limpio.
    directo = _validar_xml_factura(texto)
    if directo:
        return directo

    texto_limpio = _limpiar_texto_xml(texto)

    # 2) Buscar por regex en todo el texto original/escapado.
    inner = _extraer_xml_por_regex(texto_limpio)
    if inner:
        valido = _validar_xml_factura(inner)
        if valido:
            return valido

    # 3) Parsear AttachedDocument y revisar Description.
    try:
        root = ET.fromstring(texto_limpio)
    except Exception:
        # Si no parsea como XML, revisar si venía escapado/base64 en bruto.
        txt_unescaped = unescape(texto_limpio)
        inner = _extraer_xml_por_regex(txt_unescaped)
        if inner:
            valido = _validar_xml_factura(inner)
            if valido:
                return valido

        if _parece_base64(texto_limpio):
            decoded = _decode_base64_a_texto(texto_limpio)
            inner = _extraer_xml_por_regex(decoded)
            if inner:
                valido = _validar_xml_factura(inner)
                if valido:
                    return valido

        return None

    # Si el root ya es factura, retornarlo.
    if _local_name(root.tag).lower() in {"invoice", "creditnote", "debitnote"}:
        return texto_limpio

    for desc in _buscar_descripciones_xml(root):
        # 3.1 Texto escapado o CDATA directo.
        desc_unescaped = unescape(desc)
        inner = _extraer_xml_por_regex(desc_unescaped)
        if inner:
            valido = _validar_xml_factura(inner)
            if valido:
                return valido

        # 3.2 Base64 dentro de Description.
        if _parece_base64(desc):
            decoded = _decode_base64_a_texto(desc)
            inner = _extraer_xml_por_regex(decoded)
            if inner:
                valido = _validar_xml_factura(inner)
                if valido:
                    return valido

    return None
