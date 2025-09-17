# utils/safe_io.py
import os
import time
import random
import shutil
import pandas as pd
from typing import Optional, Iterable

def _atomic_rename(src: str, dst: str, retries: int = 5, delay: float = 0.2) -> None:
    """
    Renombra src -> dst de forma atómica con reintentos.
    Si dst existe, se reemplaza.
    """
    for i in range(retries):
        try:
            if os.path.exists(dst):
                # Reemplazo seguro en Windows
                tmp_bak = dst + f".bak_{int(time.time())}"
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
    # último intento explícito
    os.replace(src, dst)

def _cleanup_stale_tmps(final_path: str) -> None:
    """
    Limpia temporales viejos que pudieron quedar (.tmp_*, .xlsx.tmp_*, etc.).
    """
    folder = os.path.dirname(final_path) or "."
    base   = os.path.basename(final_path)

    # patrones heredados que vimos en tu carpeta
    candidates: Iterable[str] = []
    for name in os.listdir(folder):
        if not name.startswith(os.path.splitext(base)[0]):
            continue
        if (".tmp_" in name) or name.endswith(".tmp") or name.endswith(".xlsx.tmp"):
            candidates.append(os.path.join(folder, name))

    for p in candidates:
        try:
            os.remove(p)
        except OSError:
            pass

def safe_save_pandas(
    df_or_writer_input,
    final_path: str,
    sheet_name: Optional[str] = None,
    mode: str = "w",
    header: bool = True,
    index: bool = False,
):
    """
    Escribe un Excel de forma segura:
    1) Crea un **temporal que termine en .xlsx** (p.ej. facturas.tmp_abcd1234.xlsx)
    2) Escribe allí
    3) Renombra atómicamente al definitivo

    df_or_writer_input puede ser un DataFrame o un dict:
      { "dataframe": df, "writer_args": {...} }  (si quieres pasar engine_kwargs)
    """
    os.makedirs(os.path.dirname(final_path) or ".", exist_ok=True)

    _cleanup_stale_tmps(final_path)

    base, ext = os.path.splitext(final_path)
    if ext.lower() != ".xlsx":
        # por seguridad, forzamos extensión válida
        final_path = base + ".xlsx"
        ext = ".xlsx"

    # temporal con extensión .xlsx al final (clave para pandas)
    rand = f"{int(time.time())}_{random.randint(1000, 999999)}"
    tmp_path = f"{base}.tmp_{rand}{ext}"

    # construir writer
    if isinstance(df_or_writer_input, dict):
        df = df_or_writer_input.get("dataframe")
        writer_args = df_or_writer_input.get("writer_args", {})
    else:
        df = df_or_writer_input
        writer_args = {}

    # Escribir al temporal
    with pd.ExcelWriter(tmp_path, engine="openpyxl", mode="w", **writer_args) as writer:
        if sheet_name:
            df.to_excel(writer, sheet_name=sheet_name, index=index, header=header)
        else:
            df.to_excel(writer, index=index, header=header)

    # Renombrar atómicamente
    _atomic_rename(tmp_path, final_path)
    return final_path
