# utils/text_normalizer.py
import re
import unicodedata

def normalize_text(s: str) -> str:
    """
    Normaliza texto para comparaciones:
    - lowercase
    - sin tildes/acentos
    - espacios colapsados
    """
    if not s:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip()
    return s
