# services/utils/pdf_utils.py
import re
from typing import Optional, Dict


def extraer_texto_pdf(local_pdf_path: str) -> str:
    """
    Extrae texto de un PDF 'searchable'. Requiere pdfminer.six:
      pip install pdfminer.six
    Si falla, retorna cadena vacía (no rompemos el flujo).
    """
    try:
        from pdfminer.high_level import extract_text  # nombre estándar del paquete
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
    - Devuelve la cadena normalizada.
    """
    if not s:
        return ""
    return s.translate(_LIGATURE_MAP)


def _clean_hex_chunks(s: str) -> str:
    """
    Toma una cadena con posibles espacios/guiones/saltos y deja solo [0-9a-f],
    en minúsculas. Útil para CUFE/UUID partido visualmente en el PDF.
    """
    s = re.sub(r"[^0-9a-fA-F]", "", s)
    return s.lower()


# --- Regex base ---
_RE_CUFE_SIMPLE = re.compile(
    r"(CUFE|CUFD|UUID)\s*[:=]?\s*([A-Za-z0-9\-]{20,})",
    re.IGNORECASE,
)

# Variantes comunes para número (mejorado para Factura Electrónica)
_RE_NUM = re.compile(
    r"(?:Factura\s*(?:Electr[oó]nica\s*de\s*Venta)?\s*[:#]?\s*|N[o°\.]?\s*|N[úu]mero\s*[:#]?\s*)([A-Za-z0-9\-\/\.]{3,})",
    re.IGNORECASE,
)

# Fechas típicas
_RE_FEC1 = re.compile(r"(\d{4}[-/]\d{2}[-/]\d{2})")
_RE_FEC2 = re.compile(r"(\d{2}[-/]\d{2}[-/]\d{4})")


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

    # --- 4) Número de factura ---
    for match in _RE_NUM.finditer(texto):
        num = match.group(1).strip()
        if "NIT" in num.upper() or "PANADERIA" in num.upper() or "JOYCO" in num.upper():
            continue
        if 3 <= len(num) <= 40:
            out.setdefault("NUMERO", num)
            break

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
