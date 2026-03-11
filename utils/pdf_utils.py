# utils/pdf_utils.py
import re
import csv
from pathlib import Path
from typing import Optional, Dict, List

# -----------------------------
# PDF text extraction
# -----------------------------
def extraer_texto_pdf(local_pdf_path: str) -> str:
    """
    Extrae texto de un PDF searchable.
    Requiere pdfminer.six:
        pip install pdfminer.six
    Si falla, retorna cadena vacía para no romper el flujo.
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
# Normalización básica
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
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s.strip()

def _clean_spaces(s: str) -> str:
    s = (s or "").replace("\r", "\n").replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def _clean_hex_chunks(s: str) -> str:
    return re.sub(r"[^0-9a-fA-F]", "", s or "").lower()


# -----------------------------
# Fechas
# -----------------------------
_RE_FEC1 = re.compile(r"\b(\d{4}[-/]\d{2}[-/]\d{2})\b")
_RE_FEC2 = re.compile(r"\b(\d{2}[-/]\d{2}[-/]\d{4})\b")

def normalizar_fecha(fecha_str: str) -> Optional[str]:
    """
    Convierte a YYYY-MM-DD si es posible.
    Acepta:
    - YYYY/MM/DD
    - YYYY-MM-DD
    - DD/MM/YYYY
    - DD-MM-YYYY
    """
    try:
        import datetime as dt

        s = (fecha_str or "").strip()
        if not s:
            return None

        s = s.replace("\\", "/").replace(".", "/").replace("-", "/")
        parts = s.split("/")
        if len(parts) != 3:
            return None

        if len(parts[0]) == 4:
            y, m, d = map(int, parts)
        else:
            d, m, y = map(int, parts)

        return dt.date(y, m, d).strftime("%Y-%m-%d")
    except Exception:
        return None


# -----------------------------
# Reglas y filtros para número
# -----------------------------
_FACT_PREFIXES = (
    "FEAC", "FETR", "FVE", "FVN", "FE", "FV", "FCII", "FC", "FD",
    "NC", "ND", "RH", "MQ", "GS", "EN", "BP", "CO", "SKYB",
    "HYSB", "FHM", "EB", "N", "A", "H4Z", "FPP", "FEC", "FET"
)

_BAD_NUM_WORDS = {
    "CALLE", "CARRERA", "CR", "CRA", "AV", "AVENIDA", "TRANSV", "TRANSVERSAL",
    "DIAGONAL", "DG", "KM", "VIA", "PROYECTO", "PROY", "NOTARIA", "CIIU",
    "MUNICIPIO", "CIUDAD", "NIT", "CUFE", "CUFD", "UUID", "CONTRATO",
    "REFERENCIA", "PAGO", "CLIENTE", "EMISOR", "ADQUIRIENTE", "FACTURACION",
    "FACTURACIÓN", "RESOLUCION", "RESOLUCIÓN", "ACTIVIDAD", "ECONOMICA",
    "ECONÓMICA", "DIRECCION", "DIRECCIÓN", "CODIGO", "CÓDIGO", "PAIS",
    "PAÍS", "DEPARTAMENTO", "TELEFONO", "TELÉFONO", "CORREO", "BARRIO"
}

def _clean_candidate(raw: str) -> str:
    raw = (raw or "").strip()
    raw = raw.replace("–", "-").replace("—", "-")
    raw = re.sub(r"\s+", " ", raw)
    raw = re.sub(r"\s*-\s*", "-", raw)
    return raw.strip(" -")

def _num_digits(s: str) -> int:
    return len(re.findall(r"\d", s or ""))

def _has_letters(s: str) -> bool:
    return bool(re.search(r"[A-Za-z]", s or ""))

def _is_bad_num_candidate(s: str) -> bool:
    s = _clean_candidate(s).upper()
    if not s:
        return True

    if len(s) < 3:
        return True

    if s in _BAD_NUM_WORDS:
        return True

    parts = re.split(r"[-\s/]+", s)
    if parts and parts[0] in _BAD_NUM_WORDS:
        return True

    if s.startswith(("CALLE ", "CARRERA ", "CR ", "CRA ", "AV ", "AVENIDA ", "TRANSV ")):
        return True

    if s.startswith(("CIIU ", "NOTARIA ", "PROYECTO ", "NIT ")):
        return True

    # Solo texto sin números -> no sirve
    if not re.search(r"\d", s):
        return True

    # Caso tipo "FEB-26" o "MARZO 2026" o similares
    if re.fullmatch(r"[A-Z]{2,10}-\d{2,4}", s):
        monthish = {"ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"}
        pref = s.split("-")[0]
        if pref in monthish:
            return True

    if re.fullmatch(r"(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)\s+\d{4}", s):
        return True

    # "y 100", "S 199", etc.
    if re.fullmatch(r"[A-ZY]\s+\d{2,6}", s):
        return True

    return False

def _normalize_numero_factura(s: str) -> str:
    s = _clean_candidate(s)

    # unir prefijo+numero si viene separado: FE 1985 -> FE1985
    m = re.fullmatch(r"([A-Za-z]{1,10})\s+(\d{2,20})", s)
    if m:
        pref = m.group(1).upper()
        num = m.group(2)
        if pref in _FACT_PREFIXES or len(pref) <= 5:
            return f"{pref}{num}"

    # si ya trae guion, mantener pero sin espacios
    m = re.fullmatch(r"([A-Za-z]{1,10})-(\d{2,20})", s)
    if m:
        return f"{m.group(1).upper()}{m.group(2)}"

    return s

def _candidate_score(num: str, full_text: str = "") -> int:
    n = _clean_candidate(num)
    up = n.upper()
    score = 0

    if _is_bad_num_candidate(up):
        return -10_000

    if any(up.startswith(p) for p in _FACT_PREFIXES):
        score += 200

    if "-" in n:
        score += 30

    if _num_digits(n) >= 4:
        score += 25

    if _has_letters(n):
        score += 20

    # Penalizar si parece solo número demasiado genérico pero muy largo
    if re.fullmatch(r"\d{10,25}", n):
        score -= 30

    # Mejor si aparece cerca de palabra factura
    if full_text:
        pat = re.escape(n).replace(r"\ ", r"\s+")
        if re.search(rf"(factura|número de factura|numero de factura).{{0,80}}{pat}", full_text, re.IGNORECASE | re.DOTALL):
            score += 120

    return score


# -----------------------------
# CUFE
# -----------------------------
_RE_CUFE_SIMPLE = re.compile(
    r"\b(?:CUFE|CUFD|UUID)\b\s*[:=]?\s*([0-9a-fA-F\-\s]{20,})",
    re.IGNORECASE,
)

def _extraer_cufe_cercano_a_label(texto: str) -> Optional[str]:
    if not texto:
        return None

    m = re.search(r"\b(CUFE|UUID|CUFD)\b", texto, flags=re.IGNORECASE)
    if not m:
        return None

    after = texto[m.end(): m.end() + 1200]
    mhex = re.search(r"([0-9a-fA-F][0-9a-fA-F\s\-]{60,220})", after)
    if not mhex:
        return None

    cufe = _clean_hex_chunks(mhex.group(1))
    if len(cufe) >= 96:
        return cufe[:96]
    if len(cufe) >= 64:
        return cufe
    return None


# -----------------------------
# Número de factura
# -----------------------------
_RE_NUM_LABEL_1 = re.compile(
    r"(?:n[uú]mero\s+de\s+factura|numero\s+de\s+factura|factura(?:\s+electr[oó]nica)?(?:\s+de\s+venta)?\s*(?:no\.?|nro\.?|n[°ºo]|n[uú]mero|numero)?)\s*[:#]?\s*([A-Z0-9][A-Z0-9\-/ ]{1,30}\d{1,20})",
    re.IGNORECASE,
)

_RE_NUM_LABEL_2 = re.compile(
    r"\bFactura\s*[:#]?\s*([A-Z0-9][A-Z0-9\-/ ]{1,30}\d{1,20})",
    re.IGNORECASE,
)

_RE_PREFIXED = re.compile(
    rf"\b({'|'.join(sorted(_FACT_PREFIXES, key=len, reverse=True))})\s*[- ]?\s*(\d{{2,20}})\b",
    re.IGNORECASE,
)

_RE_ALNUM_COMPACT = re.compile(r"\b([A-Z]{1,8}\d{3,20})\b", re.IGNORECASE)
_RE_ALNUM_SPACED = re.compile(r"\b([A-Z]{1,8}\s+\d{2,20})\b", re.IGNORECASE)
_RE_ALNUM_HYPHEN = re.compile(r"\b([A-Z]{1,8}-\d{2,20})\b", re.IGNORECASE)
_RE_NUM_PLAIN_LONG = re.compile(r"\b(\d{6,25})\b")


def _pick_best_numero(texto: str) -> Optional[str]:
    if not texto:
        return None

    text = _normalize_text(texto)
    candidates: List[str] = []

    # 1) Etiquetas explícitas
    for rx in (_RE_NUM_LABEL_1, _RE_NUM_LABEL_2):
        for m in rx.finditer(text):
            cand = _clean_candidate(m.group(1))
            if cand:
                candidates.append(cand)

    # 2) Prefijos fuertes conocidos
    for m in _RE_PREFIXED.finditer(text):
        pref = m.group(1).upper()
        num = m.group(2)
        candidates.append(f"{pref}{num}")

    # 3) Compactos, espaciados, con guion
    for rx in (_RE_ALNUM_HYPHEN, _RE_ALNUM_SPACED, _RE_ALNUM_COMPACT):
        for m in rx.finditer(text):
            cand = _clean_candidate(m.group(1))
            if cand:
                candidates.append(cand)

    # 4) Solo número largo como último recurso
    for m in _RE_NUM_PLAIN_LONG.finditer(text):
        cand = _clean_candidate(m.group(1))
        if cand:
            candidates.append(cand)

    # Dedup manteniendo orden
    seen = set()
    uniq = []
    for c in candidates:
        key = c.upper()
        if key in seen:
            continue
        seen.add(key)
        uniq.append(c)

    best = None
    best_score = -10_000

    for cand in uniq:
        norm = _normalize_numero_factura(cand)
        score = _candidate_score(norm, text)
        if score > best_score:
            best_score = score
            best = norm

    return best if best_score > 0 else None


# -----------------------------
# Número alterno / aprobación
# -----------------------------
_RE_CONTRATO = re.compile(r"\bContrato\b[^0-9A-Za-z]{0,20}([A-Z0-9\-]{4,30})", re.IGNORECASE)
_RE_REF_PAGO = re.compile(r"\bReferencia\s+de\s+pago\b[^0-9A-Za-z]{0,20}([A-Z0-9\-]{4,30})", re.IGNORECASE)
_RE_PAGA_NUM = re.compile(r"\bPaga\s+con\s+este\s+n[uú]mero\b[^0-9A-Za-z]{0,20}([A-Z0-9\-]{4,30})", re.IGNORECASE)

def _extraer_numero_aprobacion(texto: str, numero_principal: str = "") -> Optional[str]:
    if not texto:
        return None

    for rx in (_RE_CONTRATO, _RE_REF_PAGO, _RE_PAGA_NUM):
        m = rx.search(texto)
        if m:
            cand = _clean_candidate(m.group(1))
            if cand and cand != numero_principal and not _is_bad_num_candidate(cand):
                return _normalize_numero_factura(cand)

    return None


# -----------------------------
# Fecha
# -----------------------------
def _extraer_fecha_emision(texto: str) -> Optional[str]:
    patrones = [
        r"Fecha\s+de\s+Emisi[oó]n\s*[:#]?\s*([0-9]{2}[\/\-][0-9]{2}[\/\-][0-9]{4}|[0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})",
        r"Fecha\s+de\s+factura\s*[:#]?\s*([0-9]{2}[\/\-][0-9]{2}[\/\-][0-9]{4}|[0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})",
        r"Fecha\s*[:#]?\s*([0-9]{2}[\/\-][0-9]{2}[\/\-][0-9]{4}|[0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})",
    ]

    for pat in patrones:
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            f = normalizar_fecha(m.group(1))
            if f:
                return f

    m1 = _RE_FEC1.search(texto)
    if m1:
        return normalizar_fecha(m1.group(1))

    m2 = _RE_FEC2.search(texto)
    if m2:
        return normalizar_fecha(m2.group(1))

    return None


# -----------------------------
# Parse principal
# -----------------------------
def parse_identificadores_pdf(texto: str) -> Dict[str, str]:
    """
    Extrae principalmente:
      - CUFE
      - NUMERO
      - NUMERO_APROB
      - FECHA

    Prioriza:
      1. CUFE etiquetado o hex fuerte
      2. Número real de factura
      3. Reduce falsos positivos
    """
    out: Dict[str, str] = {}
    texto = _normalize_text(texto or "")

    # CUFE
    cufe = _extraer_cufe_cercano_a_label(texto)
    if not cufe:
        m = _RE_CUFE_SIMPLE.search(texto)
        if m:
            cufe = _clean_hex_chunks(m.group(1))
            if len(cufe) >= 96:
                cufe = cufe[:96]
            elif len(cufe) < 64:
                cufe = None

    if not cufe:
        flat = _clean_hex_chunks(texto)
        m = re.search(r"([0-9a-f]{96})", flat)
        if m:
            cufe = m.group(1)

    if cufe:
        out["CUFE"] = cufe

    # Número
    numero = _pick_best_numero(texto)
    if numero:
        out["NUMERO"] = numero

    # Número alterno / aprobación
    num_aprob = _extraer_numero_aprobacion(texto, numero_principal=numero or "")
    if num_aprob:
        out["NUMERO_APROB"] = num_aprob

    # Fecha
    fecha = _extraer_fecha_emision(texto)
    if fecha:
        out["FECHA"] = fecha

    print("\n===== DEBUG PDF PARSE =====")
    print(f"→ CUFE detectado: {out.get('CUFE')}")
    print(f"→ NUMERO detectado: {out.get('NUMERO')}")
    print(f"→ NUMERO_APROB detectado: {out.get('NUMERO_APROB')}")
    print(f"→ FECHA detectada: {out.get('FECHA')}")
    print("===========================\n")

    return out


# --------------------------------------------------------
# Códigos de ciudad desde CSV externo
# --------------------------------------------------------
def _strip_accents_upper(s: str) -> str:
    rep = str.maketrans("ÁÉÍÓÚÜÑáéíóúüñ", "AEIOUUNaeiouun")
    return (s or "").translate(rep).upper().strip()

def _norm_city_key(s: str) -> str:
    s = _strip_accents_upper(s)
    s = s.replace(".", "").replace(",", "").replace(";", "").replace(":", "")
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\bD\s*C\b", "DC", s)
    return s

def _cargar_codigos_ciudad() -> Dict[str, str]:
    candidates = [
        Path("data") / "codigos_ciudad.csv",
        Path("codigos_ciudad.csv"),
    ]

    for p in candidates:
        if not (p.exists() and p.is_file()):
            continue

        try:
            raw = p.read_text(encoding="utf-8", errors="ignore")
            if not raw.strip():
                return {}

            sample = raw[:4096]
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;|\t")
            except Exception:
                dialect = csv.excel
                dialect.delimiter = ","

            reader = csv.reader(raw.splitlines(), dialect)
            mapping: Dict[str, str] = {}

            header = None
            for row in reader:
                if not row:
                    continue

                row = [c.strip() for c in row if c is not None]
                if not row or (len(row) == 1 and not row[0]):
                    continue

                if header is None:
                    low = ",".join(row).lower()
                    if "ciudad" in low and "codigo" in low:
                        header = [c.strip().lower() for c in row]
                        continue
                    header = []

                if header:
                    def idx(colname: str) -> int:
                        try:
                            return header.index(colname)
                        except Exception:
                            return -1

                    i_city = idx("ciudad")
                    i_code = idx("codigo")
                    if i_city < 0 or i_code < 0:
                        i_city, i_code = 0, 1
                    if len(row) <= max(i_city, i_code):
                        continue
                    city_raw = row[i_city]
                    code = row[i_code]
                else:
                    if len(row) < 2:
                        continue
                    city_raw = row[0]
                    code = row[1]

                city_key = _norm_city_key(city_raw)
                if city_key and code:
                    mapping[city_key] = code.strip()
                    if city_key.endswith(" DC"):
                        mapping.setdefault(city_key.replace(" DC", ""), code.strip())

            return mapping
        except Exception:
            return {}

    return {}

_CITY_CODES = None

def _codigo_ciudad(nombre_ciudad: str) -> str:
    global _CITY_CODES
    if _CITY_CODES is None:
        _CITY_CODES = _cargar_codigos_ciudad()

    key = _norm_city_key(nombre_ciudad)
    if not key:
        return ""
    return _CITY_CODES.get(key) or ""


# --------------------------------------------------------
# Descripción items
# --------------------------------------------------------
def extraer_descripcion_items_pdf(texto: str) -> str:
    t = _clean_spaces(texto)
    lines = [ln.strip() for ln in t.split("\n") if ln.strip()]

    def find_idx(pat: str, start: int = 0) -> int:
        for i in range(start, len(lines)):
            if re.search(pat, lines[i], flags=re.IGNORECASE):
                return i
        return -1

    i_det = find_idx(r"Detalles\s+de\s+Productos")
    if i_det < 0:
        return ""

    i_end = find_idx(r"(Datos\s+Totales|Notas\s+Finales|CUFE|CUDS|C[oó]digo\s+QR)", start=i_det + 1)
    if i_end < 0:
        i_end = min(len(lines), i_det + 200)

    seg = lines[i_det:i_end]

    i_nro = -1
    for i, ln in enumerate(seg):
        if re.fullmatch(r"Nro\.?", ln, flags=re.IGNORECASE):
            i_nro = i
            break
    if i_nro >= 0:
        seg = seg[i_nro + 1:]

    header_stop = set(map(str.lower, [
        "código", "codigo", "descripción", "descripcion", "u/m", "cantidad",
        "precio unitario", "subtotal", "valor total", "impuestos", "total",
        "iva %", "inc %", "dcto detalle"
    ]))

    def is_item_no(s: str) -> bool:
        return bool(re.fullmatch(r"\d{1,3}", s or ""))

    def looks_numeric(s: str) -> bool:
        return bool(re.fullmatch(r"[\d\.,\-]+", s or ""))

    def is_unit(s: str) -> bool:
        return bool(re.fullmatch(r"(UN|UND|KG|LT|GL|NIU|EA|H87|94|ZZ|GAL|LTS|LTR)", (s or "").strip().upper()))

    descs: List[str] = []
    i = 0
    while i < len(seg):
        ln = seg[i]
        if is_item_no(ln):
            i += 1
            if i < len(seg) and re.fullmatch(r"\d{1,20}", seg[i]):
                i += 1

            parts: List[str] = []
            while i < len(seg):
                cur = seg[i].strip()
                cur_low = cur.lower()

                if is_item_no(cur):
                    break
                if cur_low in header_stop:
                    break
                if looks_numeric(cur) or is_unit(cur):
                    break
                if re.search(r"\bIVA\b|\bRETENCI", cur, flags=re.IGNORECASE):
                    break

                parts.append(cur)
                i += 1

            desc = re.sub(r"\s{2,}", " ", " ".join(parts)).strip()
            if desc:
                descs.append(desc)
        else:
            i += 1

    seen = set()
    out = []
    for d in descs:
        k = d.lower()
        if k in seen:
            continue
        seen.add(k)
        out.append(d)

    return "; ".join(out).strip()


# --------------------------------------------------------
# Cabecera desde PDF
# --------------------------------------------------------
def extraer_campos_basicos_pdf(texto: str) -> Dict[str, str]:
    t = _clean_spaces(texto)
    lines = [ln.strip() for ln in t.split("\n") if ln.strip()]

    def idx_of(pat: str) -> int:
        for i, ln in enumerate(lines):
            if re.search(pat, ln, flags=re.IGNORECASE):
                return i
        return -1

    i_em = idx_of(r"Datos\s+del\s+Emisor")
    i_ad = idx_of(r"Datos\s+del\s+Adquiriente")

    if i_em < 0:
        i_em = 0
    if i_ad < 0:
        i_ad = len(lines)

    em_lines = lines[i_em:i_ad]
    ad_lines = lines[i_ad:i_ad + 250]

    def find_after(prefix_pat: str, arr) -> str:
        for ln in arr:
            m = re.search(prefix_pat, ln, flags=re.IGNORECASE)
            if m:
                return (m.group(1) or "").strip()
        return ""

    empresa = find_after(r"Raz[oó]n\s+Social:\s*(.+)", em_lines).strip()
    if empresa:
        empresa = re.split(r"Nombre\s+Comercial\s*:", empresa, flags=re.IGNORECASE)[0].strip()

    nit = find_after(r"Nit\s+del\s+Emisor:\s*([0-9\.\-]+)", em_lines)
    nit = re.sub(r"[^\d]", "", nit)

    ciudad = find_after(r"Municipio\s*/\s*Ciudad:\s*(.+)", em_lines).strip()
    if not ciudad:
        m = re.search(r"Municipio\s*/\s*Ciudad:\s*(.+)", t, flags=re.IGNORECASE)
        if m:
            ciudad = (m.group(1) or "").strip()

    regimen = find_after(r"R[eé]gimen\s+Fiscal:\s*(.+)", em_lines).strip()
    tipo_txt = find_after(r"Tipo\s+de\s+Contribuyente:\s*(.+)", em_lines).strip()
    tipo_out = regimen or tipo_txt

    act = find_after(r"Actividad\s+Econ[oó]mica:\s*([0-9;\s]+)", em_lines)
    act = re.sub(r"\s+", "", act)

    cliente = find_after(r"Nombre\s+o\s+Raz[oó]n\s+Social:\s*(.+)", ad_lines).strip()

    desc_items = extraer_descripcion_items_pdf(texto)

    return {
        "Empresa emisora": empresa,
        "Ciudad emisora": (ciudad or "").upper(),
        "Código ciudad": _codigo_ciudad(ciudad),
        "NIT": nit,
        "Cliente": cliente,
        "Tipo de contribuyente": tipo_out,
        "Actividad económica": act,
        "DescripcionLineas": desc_items,
    }


# --------------------------------------------------------
# Totales
# --------------------------------------------------------
_MONEY = r"(\d{1,3}(?:[.\s]\d{3})*(?:[.,]\d{2})|\d+(?:[.,]\d{2})?)"

def _to_float_money(s: str) -> float:
    s = (s or "").strip()
    if not s:
        return 0.0

    s = s.replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            if s.count(".") > 1:
                s = s.replace(".", "")

    try:
        return float(s)
    except Exception:
        return 0.0

def _extraer_totales_datos_totales_dian(texto: str) -> Dict[str, float]:
    t = _clean_spaces(texto)

    m = re.search(r"\nCOP\s*\n", t)
    if not m:
        return {}

    tail = t[m.end():]
    vals = re.findall(_MONEY, tail)
    if len(vals) < 13:
        return {}

    vals = vals[:13]
    nums = [_to_float_money(v) for v in vals]

    return {
        "Subtotal": float(nums[0]),
        "IVA": float(nums[4]),
        "Total": float(nums[12]),
        "Total neto": float(nums[9]),
        "Total impuesto": float(nums[8]),
        "Total bruto": float(nums[3]),
    }

def extraer_totales_basicos_pdf(texto: str) -> Dict[str, float]:
    dian = _extraer_totales_datos_totales_dian(texto)
    if dian:
        iva_total = dian.get("IVA", 0.0)
        return {
            "Subtotal": float(dian.get("Subtotal", 0.0) or 0.0),
            "IVA 19%": float(iva_total or 0.0),
            "IVA 5%": 0.0,
            "Total": float(dian.get("Total", 0.0) or 0.0),
        }

    t = _clean_spaces(texto)
    low = t.lower()

    def pick(patterns) -> float:
        for pat in patterns:
            m = re.search(pat, low, flags=re.IGNORECASE | re.DOTALL)
            if m:
                return _to_float_money(m.group(1))
        return 0.0

    subtotal = pick([
        rf"\bsubtotal\b.*?{_MONEY}",
        rf"\bbase\b.*?{_MONEY}",
    ])

    iva19 = pick([
        rf"\biva\b.*?19%.*?{_MONEY}",
        rf"\b19%\b.*?{_MONEY}",
    ])

    iva5 = pick([
        rf"\biva\b.*?5%.*?{_MONEY}",
        rf"\b5%\b.*?{_MONEY}",
    ])

    total = pick([
        rf"\btotal\s+neto\s+factura\b.*?{_MONEY}",
        rf"\btotal\s+factura\b.*?{_MONEY}",
        rf"\btotal\b.*?a\s*pagar\b.*?{_MONEY}",
        rf"\bvalor\s+total\b.*?{_MONEY}",
    ])

    return {
        "Subtotal": float(subtotal or 0.0),
        "IVA 19%": float(iva19 or 0.0),
        "IVA 5%": float(iva5 or 0.0),
        "Total": float(total or 0.0),
    }