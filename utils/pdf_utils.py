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
        from pdfminer.high_level import extract_text
        return extract_text(local_pdf_path) or ""
    except Exception as e:
        print(f"[PDF] No se pudo extraer texto: {e}")
        return ""

# -----------------------------
# Utilidades internas
# -----------------------------
def _clean_hex_chunks(s: str) -> str:
    """
    Toma una cadena con posibles espacios/guiones/saltos y deja solo [0-9a-f],
    en minúsculas. Útil para CUFE/UUID partido visualmente en el PDF.
    """
    s = re.sub(r'[^0-9a-fA-F]', '', s)
    return s.lower()

# --- Regex base (se siguen usando) ---
# Nota: dejamos los patrones existentes y añadimos una ruta más robusta
# para CUFE/UUID sin romper lo que ya funcionaba.
_RE_CUFE_SIMPLE = re.compile(r"(CUFE|CUFD|UUID)\s*[:=]?\s*([A-Za-z0-9\-]{20,})", re.IGNORECASE)

# Variantes comunes para número (No., N°, Numero, Factura)
_RE_NUM  = re.compile(
    r"(N[o°\.]?\s*|Numero\s*|Factura\s*#?\s*)([:=]?\s*)([A-Za-z0-9\-\/\.]{3,})",
    re.IGNORECASE
)

# Fechas típicas YYYY-MM-DD o DD/MM/YYYY o DD-MM-YYYY
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
            y, m, d = map(int, parts)         # YYYY/MM/DD
        else:
            d, m, y = map(int, parts)         # DD/MM/YYYY
        return dt.date(y, m, d).strftime("%Y-%m-%d")
    except Exception:
        return None

def parse_identificadores_pdf(texto: str) -> Dict[str, str]:
    """
    Intenta extraer CUFE (preferido) y, como respaldo, Número y Fecha.
    - Tolera variantes 'UUID (CUFE)', saltos de línea, espacios y guiones en el CUFE.
    - Mantiene los patrones previos para NUMERO/FECHA.
    Retorna dict con llaves posibles: {"CUFE": "...", "NUMERO": "...", "FECHA": "YYYY-MM-DD"}
    """
    out: Dict[str, str] = {}

    # --- 1) CUFE / UUID robusto ---
    # a) Búsqueda con etiqueta explícita y valor con huecos (60-140 nibbles con separadores)
    m = re.search(
        r'(?:UUID\s*\(?\s*CUFE\)?|UUID|CUFE)\s*[:\-]?\s*((?:[0-9a-fA-F][\s\-]?){60,140})',
        texto,
        flags=re.IGNORECASE
    )
    if m:
        cufe = _clean_hex_chunks(m.group(1))
        if len(cufe) >= 50:  # longitud razonable para CUFE
            out["CUFE"] = cufe

    # b) Si no, usa el patrón simple legado (por compatibilidad)
    if "CUFE" not in out:
        m = _RE_CUFE_SIMPLE.search(texto)
        if m:
            # Si viene con guiones, los quitamos de forma segura.
            raw = m.group(2).strip()
            cleaned = raw.replace("-", "")
            # Si parece hex largo, límpialo; si no, deja el valor original.
            cleaned_hex = _clean_hex_chunks(cleaned)
            out["CUFE"] = cleaned_hex if len(cleaned_hex) >= 50 else cleaned

    # --- 2) Número de factura (evita capturar 'NIT') ---
    for match in _RE_NUM.finditer(texto):
        num = match.group(3).strip()
        if "NIT" in num.upper():
            continue
        if 3 <= len(num) <= 40:
            out.setdefault("NUMERO", num)
            break

    # --- 3) Fecha: YYYY-MM-DD o DD/MM/YYYY ---
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

    return out
