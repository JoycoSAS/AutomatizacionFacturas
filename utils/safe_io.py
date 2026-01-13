# utils/safe_io.py
import os
import time
import random
import pandas as pd
from typing import Optional, Iterable, Union
from pathlib import Path

PathLike = Union[str, os.PathLike]

def _as_str(p: PathLike) -> str:
    return str(p) if not isinstance(p, str) else p

def _atomic_rename(src: PathLike, dst: PathLike, retries: int = 5, delay: float = 0.2) -> None:
    """
    Renombra src -> dst de forma atómica con reintentos.
    Si dst existe, se reemplaza.
    """
    src = _as_str(src)
    dst = _as_str(dst)

    for i in range(retries):
        try:
            if os.path.exists(dst):
                tmp_bak = f"{dst}.bak_{int(time.time())}"
                os.replace(dst, tmp_bak)
                try:
                    os.replace(src, dst)
                finally:
                    try:
                        os.remove(tmp_bak)
                    except OSError:
                        pass
            else:
                os.replace(src, dst)
            return
        except PermissionError:
            time.sleep(delay * (i + 1))

    os.replace(src, dst)

def _cleanup_stale_tmps(final_path: PathLike) -> None:
    """
    Limpia temporales viejos que pudieron quedar (.tmp_*, .xlsx.tmp_*, etc.).
    """
    final_path = _as_str(final_path)
    folder = os.path.dirname(final_path) or "."
    base = os.path.basename(final_path)

    candidates: Iterable[str] = []
    base_no_ext = os.path.splitext(base)[0]

    try:
        for name in os.listdir(folder):
            if not name.startswith(base_no_ext):
                continue
            if (".tmp_" in name) or name.endswith(".tmp") or name.endswith(".xlsx.tmp"):
                candidates.append(os.path.join(folder, name))
    except FileNotFoundError:
        return

    for p in candidates:
        try:
            os.remove(p)
        except OSError:
            pass

def safe_save_pandas(
    df_or_writer_input,
    final_path: PathLike,
    sheet_name: Optional[str] = None,
    mode: str = "w",
    header: bool = True,
    index: bool = False,
):
    """
    Escribe un Excel de forma segura:
    1) Crea un temporal que termine en .xlsx
    2) Escribe allí
    3) Renombra atómicamente al definitivo
    """
    final_path = _as_str(final_path)

    os.makedirs(os.path.dirname(final_path) or ".", exist_ok=True)
    _cleanup_stale_tmps(final_path)

    base, ext = os.path.splitext(final_path)
    if ext.lower() != ".xlsx":
        final_path = base + ".xlsx"
        base, ext = os.path.splitext(final_path)

    rand = f"{int(time.time())}_{random.randint(1000, 999999)}"
    tmp_path = f"{base}.tmp_{rand}{ext}"

    if isinstance(df_or_writer_input, dict):
        df = df_or_writer_input.get("dataframe")
        writer_args = df_or_writer_input.get("writer_args", {})
    else:
        df = df_or_writer_input
        writer_args = {}

    with pd.ExcelWriter(tmp_path, engine="openpyxl", mode="w", **writer_args) as writer:
        if sheet_name:
            df.to_excel(writer, sheet_name=sheet_name, index=index, header=header)
        else:
            df.to_excel(writer, index=index, header=header)

    _atomic_rename(tmp_path, final_path)
    return final_path
