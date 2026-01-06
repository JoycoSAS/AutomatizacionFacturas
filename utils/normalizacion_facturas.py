# utils/normalizacion_facturas.py

import re
import unicodedata
from typing import List


def _strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def normalizar_texto_basico(value: str) -> str:
    """
    Normaliza texto para comparaciones:
    - str + strip
    - sin tildes
    - minúsculas
    - solo a-z0-9
    """
    if value is None:
        value = ""
    s = str(value).strip()
    s = _strip_accents(s)
    s = s.lower()
    return re.sub(r"[^a-z0-9]", "", s)


def extraer_numero_factura_candidates(value: str) -> List[str]:
    """
    Genera varias "versiones" del número de factura para hacer match tolerante.
    Ejemplos de entradas reales:
      - '2025-11-12T07:11:4; FPP - 428790'
      - 'Factura: FPP - 428790'
      - 'FPP-428790'
      - 'FPP 428790'
      - '428790'
      - 'FPP428790'

    Retorna una lista de strings candidatos (sin normalizar aún).
    """
    if value is None:
        return []
    s = str(value).strip()
    if not s:
        return []

    cands = set()
    cands.add(s)

    # 1) Si viene con ';' (timestamp; factura)
    if ";" in s:
        parts = [p.strip() for p in s.split(";") if p.strip()]
        if parts:
            cands.add(parts[-1])

    # 2) Caso explícito "Factura: XXXX"
    m = re.search(r"Factura:\s*(.+)$", s, flags=re.IGNORECASE)
    if m:
        cands.add(m.group(1).strip())

    # 3) Buscar patrón tipo: PREFIJO + número (con espacios/guiones opcionales)
    #    FPP - 428790 / FE 94381 / DISL 1595 / etc.
    m = re.search(r"([A-Za-z]{1,12})\s*[-\s]*\s*(\d{3,})", s)
    if m:
        pref = m.group(1).strip()
        num = m.group(2).strip()
        cands.add(f"{pref}-{num}")
        cands.add(f"{pref} {num}")
        cands.add(f"{pref}{num}")

    # 4) Si hay un número largo suelto, también lo agregamos
    m2 = re.search(r"(\d{3,})", s)
    if m2:
        cands.add(m2.group(1))

    return [x for x in cands if x]


def claves_normalizadas_factura(value: str) -> List[str]:
    """
    Retorna claves normalizadas (solo a-z0-9) para usar como diccionario.
    Importante: devuelve varias claves, para tolerancia máxima.
    """
    out = []
    for cand in extraer_numero_factura_candidates(value):
        k = normalizar_texto_basico(cand)
        if k and k not in out:
            out.append(k)
    return out
