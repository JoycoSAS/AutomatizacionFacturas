import re
from typing import Optional, Dict


def extraer_texto_pdf(local_pdf_path: str) -> str:
    """
    Extrae texto de un PDF 'searchable'. Requiere pdfminer.six:
      pip install pdfminer.six
    Si falla, retorna cadena vacía (no rompemos el flujo).
    """
    try:
        from pdfminer.high_level import extract_text
    except Exception as e:
        print(f"[PDF] pdfminer.six no está instalado o no se pudo importar: {e}")
        return ""

    try:
        return extract_text(local_pdf_path) or ""
    except Exception as e:
        print(f"[PDF] No se pudo extraer texto: {e}")
        return ""


# -----------------------------
# Utilidades internas
# -----------------------------

_LIGATURE_MAP = {
    ord("ﬀ"): "ff",
    ord("ﬁ"): "fi",
    ord("ﬂ"): "fl",
    ord("ﬃ"): "ffi",
    ord("ﬄ"): "ffl",
    ord("ﬅ"): "st",
    ord("ﬆ"): "st",
}


def _normalize_text(s: str) -> str:
    if not s:
        return ""
    s = s.translate(_LIGATURE_MAP)
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s


def _clean_hex_chunks(s: str) -> str:
    s = re.sub(r"[^0-9a-fA-F]", "", s)
    return s.lower()


# --- Regex CUFE simple (solo si está pegado al CUFE:) ---
_RE_CUFE_SIMPLE = re.compile(
    r"(CUFE|CUFD|UUID)\s*[:=]?\s*([0-9a-fA-F\-]{20,})",
    re.IGNORECASE,
)

# Fechas típicas (genéricas)
_RE_FEC1 = re.compile(r"(\d{4}[-/]\d{2}[-/]\d{2})")
_RE_FEC2 = re.compile(r"(\d{2}[-/]\d{2}[-/]\d{4})")

# Fecha en formato "DD MM YYYY" (muy común en PDFs)
_RE_FEC3 = re.compile(r"\b(\d{2})\s+(\d{2})\s+(\d{4})\b")


# --------------------------------------------------------
# Extracción robusta de "Número de factura"
# --------------------------------------------------------

_FACT_PREFIXES = r"(?:FPP|FE|FVE|FV|FEC|FETR|FET|FC|FD|NC|ND)"

_GOOD_CTX = [
    "factura",
    "factura electrónica",
    "factura electronica",
    "factura electrónica de venta",
    "factura electronica de venta",
    "de venta",
    "factura no",
    "factura n",
    "n°",
    "no.",
    "nro",
    "numero de factura",
    "número de factura",
]

_BAD_CTX = [
    "resolución",
    "resolucion",
    "dian",
    "vigencia",
    "prefijo",
    "autorización",
    "autorizacion",
    "autoriza",
    "rango",
    "habilita",
    "numeración",
    "numeracion",
]

_BAD_HARD = ["ddi", "resol", "resolución", "resolucion"]

_RE_NUM_AFTER_LABEL = re.compile(
    r"(?:factura(?:\s+electr[oó]nica)?(?:\s+de\s+venta)?\s*(?:no\.?|nro\.?|n[°ºo]|número|numero)\s*[:#]?\s*)"
    r"([A-Z0-9]{1,10}\s*[-–—]?\s*\d{3,15})",
    re.IGNORECASE,
)

_RE_NUM_STRONG = re.compile(
    rf"(?:Factura\s*[:#]?\s*|Factura\s+No\.?\s*[:#]?\s*|N[o°º\.]?\s*[:#]?\s*|N[úu]mero\s*[:#]?\s*)?"
    rf"({_FACT_PREFIXES})\s*[-–—]?\s*(\d{{3,12}})",
    re.IGNORECASE,
)

_RE_NUM_GLUE = re.compile(rf"\b({_FACT_PREFIXES})(\d{{3,12}})\b", re.IGNORECASE)

_RE_NUM_GENERIC = re.compile(r"\b[A-Z0-9]{1,10}\s*[-–—]\s*\d{3,15}\b", re.IGNORECASE)

_RE_NUM_FALLBACK = re.compile(
    r"(?:Factura\s*(?:Electr[oó]nica\s*de\s*Venta)?\s*[:#]?\s*|"
    r"Factura\s+No\.?\s*[:#]?\s*|"
    r"N[o°º\.]?\s*[:#]?\s*|"
    r"N[úu]mero\s*[:#]?\s*)"
    r"([A-Za-z0-9]{1,12}[\s\-–—]*\d{3,15}|[A-Za-z0-9\-\/\.]{3,40})",
    re.IGNORECASE,
)


def _ctx_window(t: str, start: int, end: int, window: int = 80) -> str:
    left = t[max(0, start - window):start].lower()
    right = t[end:end + window].lower()
    return (left + " " + right).strip()


def _score_candidate(cand: str, ctx: str) -> int:
    up_cand = cand.upper()
    score = 0

    if any(k in ctx for k in _GOOD_CTX):
        score += 8

    if any(k in ctx for k in _BAD_CTX):
        score -= 10

    if up_cand.startswith("DDI-") or any(bad in ctx for bad in _BAD_HARD):
        score -= 30

    digits = re.findall(r"\d", cand)
    if len(digits) >= 5:
        score += 2
    elif len(digits) < 3:
        score -= 8

    if re.match(r"^[A-Z0-9]{1,6}\s*[-–—]\s*\d{3,15}$", cand, flags=re.I):
        score += 2

    return score


def _clean_candidate(raw: str) -> str:
    raw = (raw or "").strip()
    raw = re.sub(r"\s+", " ", raw)
    raw = raw.replace("–", "-").replace("—", "-")
    raw = re.sub(r"\s*-\s*", "-", raw)
    return raw.strip("-").strip()


def _pick_best_numero(texto: str) -> Optional[str]:
    if not texto:
        return None

    m0 = _RE_NUM_AFTER_LABEL.search(texto)
    if m0:
        cand0 = _clean_candidate(m0.group(1))
        if not cand0.upper().startswith("DDI-"):
            return cand0

    m = _RE_NUM_STRONG.search(texto)
    if m:
        pref = m.group(1).upper()
        num = m.group(2)
        return f"{pref}-{num}"

    m = _RE_NUM_GLUE.search(texto)
    if m:
        pref = m.group(1).upper()
        num = m.group(2)
        return f"{pref}-{num}"

    t = " ".join((texto or "").split())
    candidates = []
    for m3 in _RE_NUM_GENERIC.finditer(t):
        cand = _clean_candidate(m3.group(0))
        if cand.upper().startswith("DDI-"):
            continue
        ctx = _ctx_window(t, m3.start(), m3.end(), window=90)
        score = _score_candidate(cand, ctx)
        candidates.append((score, m3.start(), cand))

    if candidates:
        candidates.sort(key=lambda x: (-x[0], x[1]))
        return candidates[0][2]

    for match in _RE_NUM_FALLBACK.finditer(texto):
        cand = _clean_candidate(match.group(1) or "")
        up = cand.upper()

        if any(bad in up for bad in ["NIT", "PANADERIA", "JOYCO", "COLOMBIA"]):
            continue
        if any(bad in up for bad in ["DIAN", "RESOLUCION", "RESOLUCIÓN"]):
            continue
        if up.startswith("DDI-") or "DDI" in up:
            continue
        if not re.search(r"\d", cand):
            continue

        m2 = re.match(rf"\b({_FACT_PREFIXES})-?(\d{{3,12}})\b", cand, flags=re.I)
        if m2:
            return f"{m2.group(1).upper()}-{m2.group(2)}"

        if re.match(r"^[A-Z0-9]{1,10}-\d{3,15}$", cand, flags=re.I):
            return cand

        if len(re.findall(r"\d", cand)) >= 3:
            return cand

    return None


# --------------------------------------------------------
# ✅ NUEVO: Extraer "Contrato" / "Paga con este número" / "Referencia de pago"
# (Para facturas tipo UNE/servicios donde Aprobaciones usa el contrato)
# --------------------------------------------------------

_RE_CONTRATO = re.compile(r"\bContrato\b[^0-9]{0,30}(\d{5,15})", re.IGNORECASE)
_RE_PAGA_CON_ESTE_NUM = re.compile(r"\bPaga\s+con\s+este\s+n[uú]mero\b[^0-9]{0,30}(\d{5,15})", re.IGNORECASE)
_RE_REF_PAGO = re.compile(
    r"\bReferencia\s+de\s+pago\b[^0-9]{0,30}([0-9]{4,}(?:[-–—/][0-9A-Za-z]{2,})?)",
    re.IGNORECASE
)


def _extraer_numero_aprobacion(texto: str) -> Optional[str]:
    """
    Busca un identificador alterno para cruce con Aprobaciones:
      - Contrato 21359670
      - Paga con este número 21359670
      - Referencia de pago 9570534325-37 (si aplica)
    Retorna el candidato más "usable" (prioriza Contrato).
    """
    if not texto:
        return None

    m = _RE_CONTRATO.search(texto)
    if m:
        return (m.group(1) or "").strip()

    m = _RE_PAGA_CON_ESTE_NUM.search(texto)
    if m:
        return (m.group(1) or "").strip()

    m = _RE_REF_PAGO.search(texto)
    if m:
        return (m.group(1) or "").strip()

    return None


def normalizar_fecha(fecha_str: str) -> Optional[str]:
    """Devuelve fecha normalizada a YYYY-MM-DD si es posible."""
    try:
        import datetime as dt
        s = fecha_str.strip().replace("\\", "/").replace(".", "/").replace("-", "/")
        parts = s.split("/")
        if len(parts) != 3:
            return None
        if len(parts[0]) == 4:
            y, m, d = map(int, parts)  # YYYY/MM/DD
        else:
            d, m, y = map(int, parts)  # DD/MM/YYYY
        return dt.date(y, m, d).strftime("%Y-%m-%d")
    except Exception:
        return None


def _extraer_cufe_cercano_a_label(texto: str) -> Optional[str]:
    """
    Busca un CUFE (hex 80-120) después de la palabra CUFE/UUID
    dentro de una ventana corta.
    """
    if not texto:
        return None

    m = re.search(r"\b(CUFE|UUID)\b", texto, flags=re.IGNORECASE)
    if not m:
        return None

    after = texto[m.end(): m.end() + 600]

    mhex = re.search(r"([0-9a-fA-F][0-9a-fA-F\s\-]{70,160})", after)
    if not mhex:
        return None

    cufe = _clean_hex_chunks(mhex.group(1))

    if len(cufe) >= 96:
        return cufe[:96]

    return None


def _extraer_fecha_factura_por_label(texto: str) -> Optional[str]:
    """
    Intenta sacar la fecha de la factura cerca del label 'Fecha'
    (y evita 'vigencia').
    """
    if not texto:
        return None

    for m in re.finditer(r"\bFecha\b", texto, flags=re.IGNORECASE):
        ctx = texto[m.start(): m.start() + 200].lower()
        if "vigencia" in ctx:
            continue

        frag = texto[m.end(): m.end() + 200]

        m1 = _RE_FEC1.search(frag)
        if m1:
            return normalizar_fecha(m1.group(1))

        m2 = _RE_FEC2.search(frag)
        if m2:
            return normalizar_fecha(m2.group(1))

        m3 = _RE_FEC3.search(frag)
        if m3:
            d, mo, y = m3.group(1), m3.group(2), m3.group(3)
            return normalizar_fecha(f"{d}/{mo}/{y}")

    return None


def parse_identificadores_pdf(texto: str) -> Dict[str, str]:
    """
    Extrae CUFE (preferido) y, como respaldo, Número y Fecha.
    Además extrae NUMERO_APROB (Contrato / Paga con este número / Referencia de pago)
    para cruce con la tabla de aprobaciones.
    """
    out: Dict[str, str] = {}
    texto = _normalize_text(texto or "")

    # --- 1) CUFE cercano al label ---
    cufe_label = _extraer_cufe_cercano_a_label(texto)
    if cufe_label and len(cufe_label) == 96:
        out["CUFE"] = cufe_label

    # --- 2) CUFE simple (si venía pegado) ---
    if "CUFE" not in out:
        m = _RE_CUFE_SIMPLE.search(texto)
        if m:
            raw = m.group(2).strip()
            cleaned_hex = _clean_hex_chunks(raw)
            if len(cleaned_hex) >= 96:
                out["CUFE"] = cleaned_hex[:96]

    # --- 3) Fallback: buscar un hex largo ---
    if "CUFE" not in out:
        flat = _clean_hex_chunks(texto)
        m = re.search(r"([0-9a-f]{96})", flat)
        if m:
            out["CUFE"] = m.group(1)

    # --- 4) Número principal (normalmente el de factura/ID) ---
    numero = _pick_best_numero(texto)
    if numero:
        out["NUMERO"] = numero

    # --- 4.1) ✅ Número alterno para aprobaciones (Contrato/Ref pago) ---
    num_aprob = _extraer_numero_aprobacion(texto)
    if num_aprob:
        out["NUMERO_APROB"] = num_aprob

    # --- 5) Fecha por label ---
    fecha = _extraer_fecha_factura_por_label(texto)
    if fecha:
        out["FECHA"] = fecha

    # --- 6) Fallback genérico de fecha ---
    if "FECHA" not in out:
        m1 = _RE_FEC1.search(texto)
        if m1:
            norm = normalizar_fecha(m1.group(1))
            if norm:
                out["FECHA"] = norm

    if "FECHA" not in out:
        m2 = _RE_FEC2.search(texto)
        if m2:
            norm = normalizar_fecha(m2.group(1))
            if norm:
                out["FECHA"] = norm

    print("\n===== DEBUG PDF PARSE =====")
    print(f"→ CUFE detectado: {out.get('CUFE')}")
    print(f"→ NUMERO detectado: {out.get('NUMERO')}")
    print(f"→ NUMERO_APROB detectado: {out.get('NUMERO_APROB')}")
    print(f"→ FECHA detectada: {out.get('FECHA')}")
    print("===========================\n")

    return out
