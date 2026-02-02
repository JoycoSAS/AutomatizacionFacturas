# utils/pdf_utils.py
import re
import csv
from pathlib import Path
from typing import Optional, Dict, List, Tuple

# -----------------------------
# PDF text extraction
# -----------------------------
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
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s.strip()

def _clean_spaces(s: str) -> str:
    s = (s or "").replace("\r", "\n").replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s)
    return s.strip()

def _clean_hex_chunks(s: str) -> str:
    s = re.sub(r"[^0-9a-fA-F]", "", s)
    return s.lower()


# -----------------------------
# Fechas
# -----------------------------
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


# -----------------------------
# Identificadores (CUFE / NUMERO / NUMERO_APROB / FECHA)
# -----------------------------
_RE_CUFE_SIMPLE = re.compile(
    r"(CUFE|CUFD|UUID)\s*[:=]?\s*([0-9a-fA-F\-]{20,})",
    re.IGNORECASE,
)

_FACT_PREFIXES = r"(?:FPP|FE|FVE|FV|FEC|FETR|FET|FC|FD|NC|ND|GURA|GURC|GURD)"

_RE_NUM_AFTER_LABEL = re.compile(
    r"(?:n[uú]mero\s+de\s+factura\s*[:#]?\s*|factura(?:\s+electr[oó]nica)?(?:\s+de\s+venta)?\s*(?:no\.?|nro\.?|n[°ºo]|número|numero)\s*[:#]?\s*)"
    r"([A-Z0-9]{1,20}\s*[-–—]?\s*\d{2,20})",
    re.IGNORECASE,
)

_RE_NUM_STRONG = re.compile(
    rf"(?:Factura\s*[:#]?\s*|Factura\s+No\.?\s*[:#]?\s*|N[o°º\.]?\s*[:#]?\s*|N[úu]mero\s*[:#]?\s*)?"
    rf"({_FACT_PREFIXES})\s*[-–—]?\s*(\d{{2,20}})",
    re.IGNORECASE,
)

_RE_NUM_GLUE = re.compile(rf"\b({_FACT_PREFIXES})(\d{{2,20}})\b", re.IGNORECASE)
_RE_NUM_GENERIC = re.compile(r"\b[A-Z]{1,10}\s*[-–—]?\s*\d{2,20}\b", re.IGNORECASE)

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
        return cand0 or None

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
    for m3 in _RE_NUM_GENERIC.finditer(t):
        cand = _clean_candidate(m3.group(0))
        if cand:
            return cand

    return None

_RE_CONTRATO = re.compile(r"\bContrato\b[^0-9]{0,30}(\d{5,20})", re.IGNORECASE)
_RE_PAGA_CON_ESTE_NUM = re.compile(r"\bPaga\s+con\s+este\s+n[uú]mero\b[^0-9]{0,30}(\d{5,20})", re.IGNORECASE)
_RE_REF_PAGO = re.compile(r"\bReferencia\s+de\s+pago\b[^0-9]{0,30}([0-9]{4,}(?:[-–—/][0-9A-Za-z]{2,})?)", re.IGNORECASE)

def _extraer_numero_aprobacion(texto: str) -> Optional[str]:
    if not texto:
        return None
    for rx in (_RE_CONTRATO, _RE_PAGA_CON_ESTE_NUM, _RE_REF_PAGO):
        m = rx.search(texto)
        if m:
            return (m.group(1) or "").strip()
    return None

def _extraer_cufe_cercano_a_label(texto: str) -> Optional[str]:
    if not texto:
        return None
    m = re.search(r"\b(CUFE|UUID)\b", texto, flags=re.IGNORECASE)
    if not m:
        return None
    after = texto[m.end(): m.end() + 900]
    mhex = re.search(r"([0-9a-fA-F][0-9a-fA-F\s\-]{70,180})", after)
    if not mhex:
        return None
    cufe = _clean_hex_chunks(mhex.group(1))
    if len(cufe) >= 96:
        return cufe[:96]
    return None

def _extraer_fecha_emision(texto: str) -> Optional[str]:
    m = re.search(
        r"Fecha\s+de\s+Emisi[oó]n\s*:\s*([0-9]{2}[\/\-][0-9]{2}[\/\-][0-9]{4}|[0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})",
        texto, re.IGNORECASE
    )
    if m:
        return normalizar_fecha(m.group(1))

    m1 = _RE_FEC1.search(texto)
    if m1:
        return normalizar_fecha(m1.group(1))
    m2 = _RE_FEC2.search(texto)
    if m2:
        return normalizar_fecha(m2.group(1))
    return None

def parse_identificadores_pdf(texto: str) -> Dict[str, str]:
    """
    Extrae:
      - CUFE (preferido)
      - NUMERO (factura)
      - NUMERO_APROB (Contrato / Ref pago, etc.)
      - FECHA (YYYY-MM-DD)
    """
    out: Dict[str, str] = {}
    texto = _normalize_text(texto or "")

    cufe_label = _extraer_cufe_cercano_a_label(texto)
    if cufe_label and len(cufe_label) == 96:
        out["CUFE"] = cufe_label

    if "CUFE" not in out:
        m = _RE_CUFE_SIMPLE.search(texto)
        if m:
            raw = m.group(2).strip()
            cleaned_hex = _clean_hex_chunks(raw)
            if len(cleaned_hex) >= 96:
                out["CUFE"] = cleaned_hex[:96]

    if "CUFE" not in out:
        flat = _clean_hex_chunks(texto)
        m = re.search(r"([0-9a-f]{96})", flat)
        if m:
            out["CUFE"] = m.group(1)

    numero = _pick_best_numero(texto)
    if numero:
        out["NUMERO"] = numero

    num_aprob = _extraer_numero_aprobacion(texto)
    if num_aprob:
        out["NUMERO_APROB"] = num_aprob

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
# Códigos de ciudad desde CSV externo (robusto)
# --------------------------------------------------------
def _strip_accents_upper(s: str) -> str:
    rep = str.maketrans("ÁÉÍÓÚÜÑáéíóúüñ", "AEIOUUNaeiouun")
    return (s or "").translate(rep).upper().strip()

def _norm_city_key(s: str) -> str:
    """
    Normaliza para matching:
      - quita acentos
      - upper
      - quita . , ; :
      - colapsa espacios
      - normaliza D.C / DC / D C
    """
    s = _strip_accents_upper(s)
    s = s.replace(".", "").replace(",", "").replace(";", "").replace(":", "")
    s = re.sub(r"\s+", " ", s).strip()
    # normaliza "D C" -> "DC"
    s = re.sub(r"\bD\s*C\b", "DC", s)
    return s

def _cargar_codigos_ciudad() -> Dict[str, str]:
    """
    Soporta CSV simple en:
      - data/codigos_ciudad.csv
      - codigos_ciudad.csv

    Acepta encabezados:
      ciudad,codigo
      ciudad,codigo,depto_codigo,mun_codigo

    También acepta separador ; o , (auto-detect).
    """
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

            # detectar delimitador
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

                # detectar header
                if header is None:
                    low = ",".join(row).lower()
                    if "ciudad" in low and "codigo" in low:
                        header = [c.strip().lower() for c in row]
                        continue
                    header = []  # sin header, seguimos normal

                # si viene con header, ubicar columnas
                if header:
                    # encontrar indices
                    def idx(colname: str) -> int:
                        try:
                            return header.index(colname)
                        except Exception:
                            return -1

                    i_city = idx("ciudad")
                    i_code = idx("codigo")
                    if i_city < 0 or i_code < 0:
                        # fallback: primeras dos columnas
                        i_city, i_code = 0, 1

                    if len(row) <= max(i_city, i_code):
                        continue

                    city_raw = row[i_city]
                    code = row[i_code]
                else:
                    # sin header: primeras 2 columnas
                    if len(row) < 2:
                        continue
                    city_raw = row[0]
                    code = row[1]

                city_key = _norm_city_key(city_raw)
                if city_key and code:
                    mapping[city_key] = code.strip()

                    # alias extra: si termina en " DC", agregar variante sin DC (por si el PDF no lo trae)
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
    return (_CITY_CODES.get(key) or "")


# --------------------------------------------------------
# Descripción items desde “Detalles de Productos”
# --------------------------------------------------------
def extraer_descripcion_items_pdf(texto: str) -> str:
    """
    Extrae SOLO descripciones reales de items, evitando cabeceras/columnas numéricas.
    """
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

    # empezar DESPUÉS del header "Nro."
    i_nro = -1
    for i, ln in enumerate(seg):
        if re.fullmatch(r"Nro\.?", ln, flags=re.IGNORECASE):
            i_nro = i
            break
    if i_nro >= 0:
        seg = seg[i_nro + 1 :]

    header_stop = set(map(str.lower, [
        "código", "codigo", "descripción", "descripcion", "u/m", "cantidad",
        "precio unitario", "subtotal", "valor total", "impuestos", "total",
        "iva %", "inc %", "dcto detalle"
    ]))

    bad_prefixes = (
        "PAÍS:", "PAIS:", "DEPARTAMENTO:", "MUNICIPIO", "CIUDAD",
        "DIRECCIÓN:", "DIRECCION:", "TELÉFONO", "TELEFONO", "CORREO"
    )

    def is_item_no(s: str) -> bool:
        return bool(re.fullmatch(r"\d{1,3}", s or ""))

    def looks_numeric(s: str) -> bool:
        return bool(re.fullmatch(r"[\d\.,\-]+", s or ""))

    def is_unit(s: str) -> bool:
        return bool(re.fullmatch(r"(UN|UND|KG|LT|GL|NIU|EA|H87|94|ZZ|GAL|LTS|LTR)", (s or "").strip().upper()))

    descs = []
    i = 0
    while i < len(seg):
        ln = seg[i]

        if is_item_no(ln):
            i += 1

            # saltar código numérico si viene
            if i < len(seg) and re.fullmatch(r"\d{1,20}", seg[i]):
                i += 1

            parts = []
            while i < len(seg):
                cur = seg[i].strip()
                cur_up = cur.upper()
                cur_low = cur.lower()

                if is_item_no(cur):
                    break
                if cur_low in header_stop:
                    break
                if cur_up.startswith(bad_prefixes):
                    break
                if looks_numeric(cur) or is_unit(cur):
                    break
                if re.search(r"\bIVA\b|\bRETENCI", cur_up):
                    break

                parts.append(cur)
                i += 1

            # ✅ FIX: re.sub necesita (pattern, repl, string)
            joined = " ".join(parts)
            desc = re.sub(r"\s{2,}", " ", joined).strip()
            if desc:
                descs.append(desc)
        else:
            i += 1

    # dedup preservando orden
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
# Extracción cabecera desde PDF (DIAN)
# --------------------------------------------------------
def extraer_campos_basicos_pdf(texto: str) -> Dict[str, str]:
    """
    Extrae campos de cabecera y descripción real de items desde PDF DIAN:
      - Empresa emisora: Razón Social (sin Nombre Comercial)
      - Ciudad emisora: desde "Municipio / Ciudad:"
      - Tipo contribuyente: usa Régimen Fiscal (R-xx-xx) si existe, si no, el texto
      - Actividad económica
      - Cliente (Nombre o Razón Social del adquiriente)
      - DescripcionLineas: items reales
      - Código ciudad: lookup por CSV
    """
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
    ad_lines = lines[i_ad:i_ad + 200]

    def find_after(prefix_pat: str, arr) -> str:
        for ln in arr:
            m = re.search(prefix_pat, ln, flags=re.IGNORECASE)
            if m:
                return (m.group(1) or "").strip()
        return ""

    empresa = find_after(r"Raz[oó]n\s+Social:\s*(.+)", em_lines)
    empresa = re.split(r"Nombre\s+Comercial\s*:", empresa, flags=re.IGNORECASE)[0].strip()

    nit = find_after(r"Nit\s+del\s+Emisor:\s*([0-9\.\-]+)", em_lines)
    nit = re.sub(r"[^\d]", "", nit)

    # ✅ Ciudad: primero intenta en bloque Emisor; si no, busca en todo el texto
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
# Totales desde PDF (DIAN robusto)
# --------------------------------------------------------
_MONEY = r"(\d{1,3}(?:[.\s]\d{3})*(?:[.,]\d{2})|\d+(?:[.,]\d{2})?)"

def _to_float_money(s: str) -> float:
    s = (s or "").strip()
    if not s:
        return 0.0

    s = s.replace(" ", "")
    if "," in s and "." in s:
        # detectar decimal
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
    """
    Para PDFs DIAN como el 1100:
    - Busca el bloque donde aparece "COP" y luego una columna de 13 valores.
    - Asigna por orden oficial.
    """
    t = _clean_spaces(texto)

    # buscamos el "COP" que antecede la columna real (normalmente aparece como línea sola)
    m = re.search(r"\nCOP\s*\n", t)
    if not m:
        return {}

    tail = t[m.end():]
    vals = re.findall(_MONEY, tail)
    if len(vals) < 13:
        return {}

    # tomamos los primeros 13 valores (son los que corresponden a la columna COP)
    vals = vals[:13]
    nums = [_to_float_money(v) for v in vals]

    # orden esperado en DIAN
    # 0 Subtotal
    # 1 Descuento detalle
    # 2 Recargo detalle
    # 3 Total Bruto Factura
    # 4 IVA
    # 5 INC
    # 6 Bolsas
    # 7 Otros impuestos
    # 8 Total impuesto (=)
    # 9 Total neto factura (=)
    # 10 Descuento Global (-)
    # 11 Recargo Global (+)
    # 12 Total factura (=)
    return {
        "Subtotal": float(nums[0]),
        "IVA": float(nums[4]),
        "Total": float(nums[12]),
        "Total neto": float(nums[9]),
        "Total impuesto": float(nums[8]),
        "Total bruto": float(nums[3]),
    }

def extraer_totales_basicos_pdf(texto: str) -> Dict[str, float]:
    """
    Extrae Subtotal / IVA 19 / IVA 5 / Total.

    ✅ Primero intenta modo DIAN (tabla Datos Totales), para evitar errores como:
    - Total = Subtotal (cuando el regex agarra el primer número del bloque).
    """
    # 1) intento DIAN robusto
    dian = _extraer_totales_datos_totales_dian(texto)
    if dian:
        iva_total = dian.get("IVA", 0.0)

        # En PDFs DIAN normalmente no discrimina tarifa (5/19). Si no está explícito:
        # - Ponemos el IVA en IVA 19% (como han venido quedando tus conceptos)
        return {
            "Subtotal": float(dian.get("Subtotal", 0.0) or 0.0),
            "IVA 19%": float(iva_total or 0.0),
            "IVA 5%": 0.0,
            "Total": float(dian.get("Total", 0.0) or 0.0),
        }

    # 2) fallback genérico (si el PDF no es DIAN tabla)
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

    # IVA con porcentaje explícito
    iva19 = pick([
        rf"\biva\b.*?19%.*?{_MONEY}",
        rf"\b19%\b.*?{_MONEY}",
    ])

    iva5 = pick([
        rf"\biva\b.*?5%.*?{_MONEY}",
        rf"\b5%\b.*?{_MONEY}",
    ])

    # Total: priorizar "total neto factura" si existe
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
