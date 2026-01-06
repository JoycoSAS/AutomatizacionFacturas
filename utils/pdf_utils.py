# utils/pdf_utils.py
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

# Mapa de ligaduras típicas que aparecen en algunos PDFs (como 'ﬀ' en lugar de 'ff')
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
    """
    Normaliza el texto extraído del PDF:
    - Reemplaza ligaduras (ﬀ, ﬁ, etc.) por sus equivalentes ASCII.
    - Unifica espacios.
    """
    if not s:
        return ""
    s = s.translate(_LIGATURE_MAP)
    # Unificar saltos y múltiples espacios
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s


def _clean_hex_chunks(s: str) -> str:
    """
    Toma una cadena con posibles espacios/guiones/saltos y deja solo [0-9a-f],
    en minúsculas. Útil para CUFE/UUID partido visualmente en el PDF.
    """
    s = re.sub(r"[^0-9a-fA-F]", "", s)
    return s.lower()


# --- Regex CUFE ---
_RE_CUFE_SIMPLE = re.compile(
    r"(CUFE|CUFD|UUID)\s*[:=]?\s*([A-Za-z0-9\-]{20,})",
    re.IGNORECASE,
)

# Fechas típicas
_RE_FEC1 = re.compile(r"(\d{4}[-/]\d{2}[-/]\d{2})")
_RE_FEC2 = re.compile(r"(\d{2}[-/]\d{2}[-/]\d{4})")


# --------------------------------------------------------
# NUEVO: extracción robusta de "Número de factura"
# --------------------------------------------------------

# Prefijos comunes reales de tus facturas (ajústalo si quieres)
# (Incluye FPP por tu caso; incluye FE/FV/FVE/FEC/FETR que vi en tus ejemplos)
_FACT_PREFIXES = r"(?:FPP|FE|FVE|FV|FEC|FETR|FET|FC|FD|NC|ND)"

# 1) Patrón fuerte: prefijo + separador opcional + dígitos
# Soporta:
#   FPP - 428790
#   FPP-428790
#   FPP 428790
#   N° FPP - 428790
#   Factura: FPP - 428790
_RE_NUM_STRONG = re.compile(
    rf"(?:Factura\s*[:#]?\s*|Factura\s+No\.?\s*[:#]?\s*|N[o°º\.]?\s*[:#]?\s*|N[úu]mero\s*[:#]?\s*)?"
    rf"({_FACT_PREFIXES})\s*[-–—]?\s*(\d{{3,12}})",
    re.IGNORECASE,
)

# 2) Patrón alterno: a veces viene todo pegado (FPP428790)
_RE_NUM_GLUE = re.compile(
    rf"\b({_FACT_PREFIXES})(\d{{3,12}})\b",
    re.IGNORECASE,
)

# 3) Fallback general (tu idea original) pero más limitado:
# Evitamos capturar palabras largas sin números.
_RE_NUM_FALLBACK = re.compile(
    r"(?:Factura\s*(?:Electr[oó]nica\s*de\s*Venta)?\s*[:#]?\s*|"
    r"Factura\s+No\.?\s*[:#]?\s*|"
    r"N[o°º\.]?\s*[:#]?\s*|"
    r"N[úu]mero\s*[:#]?\s*)"
    r"([A-Za-z]{1,10}[\s\-]*\d{3,12}|[A-Za-z0-9\-\/\.]{3,40})",
    re.IGNORECASE,
)


def _pick_best_numero(texto: str) -> Optional[str]:
    """
    Devuelve el mejor candidato de número de factura, priorizando:
    - Prefijos conocidos (FPP/FE/FVE/...)
    - Prefijos pegados
    - Fallback general con filtros
    """
    if not texto:
        return None

    # A) Patrones fuertes (prefijo + dígitos)
    m = _RE_NUM_STRONG.search(texto)
    if m:
        pref = m.group(1).upper()
        num = m.group(2)
        return f"{pref}-{num}"  # normalizamos con guion (interno)

    # B) Prefijo pegado
    m = _RE_NUM_GLUE.search(texto)
    if m:
        pref = m.group(1).upper()
        num = m.group(2)
        return f"{pref}-{num}"

    # C) Fallback general con filtros duros
    for match in _RE_NUM_FALLBACK.finditer(texto):
        raw = (match.group(1) or "").strip()

        # Normalizar separadores raros
        raw_clean = re.sub(r"\s+", " ", raw).strip()

        # Filtros anti-basura (lo que te estaba pasando con "SULTORES")
        up = raw_clean.upper()
        if any(bad in up for bad in ["NIT", "PANADERIA", "JOYCO", "COLOMBIA", "DIAN"]):
            continue

        # Debe tener al menos 1 dígito; si no, lo ignoramos
        if not re.search(r"\d", raw_clean):
            continue

        # Evitar palabras muy largas (típico de texto pegado) si casi no hay dígitos
        digits = re.findall(r"\d", raw_clean)
        if len(digits) < 3:
            continue

        # Si viene tipo "FPP - 428790" o "FE 22076" lo unificamos
        m2 = re.match(rf"\b({_FACT_PREFIXES})\s*[-–—]?\s*(\d{{3,12}})\b", raw_clean, flags=re.I)
        if m2:
            return f"{m2.group(1).upper()}-{m2.group(2)}"

        return raw_clean

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


def parse_identificadores_pdf(texto: str) -> Dict[str, str]:
    """
    Intenta extraer CUFE (preferido) y, como respaldo, Número y Fecha.
    - Tolera variantes 'UUID (CUFE)', saltos de línea, espacios y guiones en el CUFE.
    - Normaliza ligaduras (como 'ﬀ') antes de buscar.
    - Ajusta CUFEs más largos automáticamente.
    Retorna dict con llaves posibles: {"CUFE": "...", "NUMERO": "...", "FECHA": "YYYY-MM-DD"}
    """
    out: Dict[str, str] = {}
    texto = _normalize_text(texto or "")

    # --- 1) CUFE / UUID robusto ---
    m = re.search(
        r"(?:UUID\s*\(?\s*CUFE\)?|UUID|CUFE)\s*[:\-]?\s*((?:[0-9a-fA-F][\s\-]?){60,140})",
        texto,
        flags=re.IGNORECASE,
    )
    if m:
        cufe = _clean_hex_chunks(m.group(1))
        if len(cufe) >= 50:
            out["CUFE"] = cufe

    # --- 2) CUFE simple ---
    if "CUFE" not in out:
        m = _RE_CUFE_SIMPLE.search(texto)
        if m:
            raw = m.group(2).strip()
            cleaned = raw.replace("-", "")
            cleaned_hex = _clean_hex_chunks(cleaned)
            out["CUFE"] = cleaned_hex if len(cleaned_hex) >= 50 else cleaned

    # --- 3) Fallback: secuencia hex larga ---
    if "CUFE" not in out:
        flat = _clean_hex_chunks(texto)
        m = re.search(r"([0-9a-f]{80,120})", flat)
        if m:
            out["CUFE"] = m.group(1)

    # --- 4) Número de factura (NUEVO robusto) ---
    numero = _pick_best_numero(texto)
    if numero:
        out["NUMERO"] = numero

    # --- 5) Fecha ---
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

    # --- 6) Ajuste final del CUFE ---
    if "CUFE" in out and len(out["CUFE"]) > 96:
        out["CUFE"] = out["CUFE"][-96:]

    print("\n===== DEBUG PDF PARSE =====")
    print(f"→ CUFE detectado: {out.get('CUFE')}")
    print(f"→ NUMERO detectado: {out.get('NUMERO')}")
    print(f"→ FECHA detectada: {out.get('FECHA')}")
    print("===========================\n")

    return out
