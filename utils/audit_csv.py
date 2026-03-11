import os
import csv
import datetime
from typing import Dict, List, Optional

def _ensure_parent(path: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)

def _is_new_file(path: str) -> bool:
    return (not os.path.exists(path)) or os.path.getsize(path) == 0

def _truncate(s: str, n: int) -> str:
    s = (s or "").strip()
    return s if len(s) <= n else (s[:n-3] + "...")

def append_csv_row(path: str, fieldnames: List[str], row: Dict):
    """
    CSV compatible con Excel (ES):
    - separador ';'
    - UTF-8 con BOM (utf-8-sig)
    - quoting para evitar que comas en subject rompan columnas
    """
    _ensure_parent(path)
    write_header = _is_new_file(path)

    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(
            f,
            fieldnames=fieldnames,
            delimiter=";",
            quoting=csv.QUOTE_MINIMAL,
            extrasaction="ignore"
        )
        if write_header:
            w.writeheader()
        w.writerow(row)

def make_daily_paths(base_dir: str, prefix: str, date: Optional[str] = None) -> str:
    if not date:
        date = datetime.datetime.now().strftime("%Y-%m-%d")
    return os.path.join(base_dir, f"{prefix}_{date}.csv")

def build_run_row(**kwargs) -> Dict:
    return dict(kwargs)

def build_detalle_row(**kwargs) -> Dict:
    # Añadimos “short” para que sea legible
    msg_id = kwargs.get("msg_id", "") or ""
    subject = kwargs.get("subject", "") or ""

    kwargs["msg_id_short"] = _truncate(msg_id, 16)
    kwargs["subject_short"] = _truncate(subject, 120)
    return dict(kwargs)