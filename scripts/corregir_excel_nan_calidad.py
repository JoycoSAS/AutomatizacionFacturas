# scripts/corregir_excel_nan_calidad.py
# Corrección puntual del Excel ya generado:
# - NO reprocesa correos
# - NO descarga adjuntos
# - NO toca SharePoint
# - Limpia valores tipo nan/None/null
# - Recalcula Estado_calidad con la lógica activa de services/excel_service.py

from __future__ import annotations

import argparse
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

import pandas as pd


def _project_root() -> Path:
    """
    Permite ejecutar el script desde:
    - raíz del proyecto
    - carpeta scripts/
    """
    here = Path(__file__).resolve()
    return here.parents[1] if here.parent.name.lower() == "scripts" else Path.cwd().resolve()


ROOT = _project_root()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


from services.excel_service import (  # noqa: E402
    ARCHIVO_EXCEL,
    COLUMNAS_VALIDAS_FINALES,
    COLUMNA_ESTADO_CALIDAD,
    _aplicar_formato_visual_facturas,
    _factura_group_key_excel_20260515,
    _limpiar_dataframe_a_formato_largo,
    _ordenar_dataframe_facturas,
    _recalcular_estado_calidad_dataframe,
)
from utils.safe_io import safe_save_pandas  # noqa: E402


VALORES_VACIOS_EXTRA = {
    "NAN",
    "NONE",
    "NULL",
    "N/A",
    "NA",
    "SIN_DATO",
    "SIN DATO",
    "NAT",
    "<NA>",
}


def _norm_empty(v: Any) -> Any:
    """
    Convierte valores basura comunes a celda vacía.
    Mantiene valores numéricos reales.
    """
    if v is None:
        return ""

    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass

    if isinstance(v, str):
        s = v.strip()
        if not s:
            return ""
        if s.upper() in VALORES_VACIOS_EXTRA:
            return ""
        return s

    return v


def _limpiar_nan_global(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_VALIDAS_FINALES)

    work = df.copy()

    # Limpieza celda por celda para evitar que "nan" llegue como texto.
    for col in work.columns:
        work[col] = work[col].map(_norm_empty)

    return work


def _estado_counts(df: pd.DataFrame) -> Counter:
    if df is None or df.empty or COLUMNA_ESTADO_CALIDAD not in df.columns:
        return Counter()

    vals = (
        df[COLUMNA_ESTADO_CALIDAD]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    return Counter(v or "VACIO" for v in vals)


def _contar_facturas(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0

    grupos = set()
    for _, row in df.iterrows():
        grupos.add(_factura_group_key_excel_20260515(row.to_dict()))
    return len(grupos)


def _conceptos_por_estado(df: pd.DataFrame) -> Dict[str, int]:
    if df is None or df.empty or COLUMNA_ESTADO_CALIDAD not in df.columns:
        return {}

    grupos_estado: Dict[Tuple[str, str], str] = {}

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        key = _factura_group_key_excel_20260515(row_dict)
        estado = str(row_dict.get(COLUMNA_ESTADO_CALIDAD, "") or "").strip().upper() or "VACIO"
        grupos_estado[key] = estado

    return dict(Counter(grupos_estado.values()))


def _contar_nan_texto(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0

    total = 0
    for col in df.columns:
        serie = df[col].fillna("").astype(str).str.strip().str.upper()
        total += int(serie.isin(VALORES_VACIOS_EXTRA).sum())
    return total


def _print_counter(title: str, counter: Counter | Dict[str, int]) -> None:
    print(title)
    if not counter:
        print("  (sin datos)")
        return
    for k in sorted(counter.keys()):
        print(f"  {k}: {counter[k]}")


def corregir_excel(dry_run: bool = False) -> None:
    excel_path = Path(ARCHIVO_EXCEL)

    if not excel_path.exists():
        raise FileNotFoundError(f"No existe el Excel: {excel_path}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = excel_path.with_name(
        f"{excel_path.stem}_BACKUP_ANTES_CORREGIR_NAN_CALIDAD_{timestamp}{excel_path.suffix}"
    )

    print("=" * 80)
    print("CORRECCIÓN PUNTUAL EXCEL - NAN + ESTADO_CALIDAD")
    print("=" * 80)
    print(f"Proyecto raíz: {ROOT}")
    print(f"Excel objetivo: {excel_path}")
    print(f"Dry-run: {dry_run}")
    print()

    df_original = pd.read_excel(excel_path, sheet_name="Facturas", engine="openpyxl")

    print("ANTES")
    print(f"  Filas: {len(df_original)}")
    print(f"  Facturas/grupos: {_contar_facturas(_limpiar_dataframe_a_formato_largo(df_original))}")
    print(f"  Celdas con texto tipo nan/null/none: {_contar_nan_texto(df_original)}")
    _print_counter("  Estado_calidad por fila:", _estado_counts(df_original))
    _print_counter(
        "  Estado_calidad por factura:",
        _conceptos_por_estado(_limpiar_dataframe_a_formato_largo(df_original)),
    )
    print()

    # 1. Limpiar basura tipo nan/null/none.
    df = _limpiar_nan_global(df_original)

    # 2. Reaplicar formato largo oficial.
    df = _limpiar_dataframe_a_formato_largo(df)

    # 3. Eliminar filas sin concepto, por seguridad.
    if "Concepto" in df.columns:
        df = df[df["Concepto"].astype(str).str.strip() != ""].copy()

    # 4. Garantizar columnas y orden.
    for col in COLUMNAS_VALIDAS_FINALES:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNAS_VALIDAS_FINALES].copy()

    # 5. Ordenar y recalcular Estado_calidad con la lógica actual.
    df = _ordenar_dataframe_facturas(df)
    df = _recalcular_estado_calidad_dataframe(df)

    # 6. Limpieza final por si el recálculo/lectura dejó restos.
    df = _limpiar_nan_global(df)
    df = _limpiar_dataframe_a_formato_largo(df)
    df = _recalcular_estado_calidad_dataframe(df)

    print("DESPUÉS PROPUESTO")
    print(f"  Filas: {len(df)}")
    print(f"  Facturas/grupos: {_contar_facturas(df)}")
    print(f"  Celdas con texto tipo nan/null/none: {_contar_nan_texto(df)}")
    _print_counter("  Estado_calidad por fila:", _estado_counts(df))
    _print_counter("  Estado_calidad por factura:", _conceptos_por_estado(df))
    print()

    if len(df) != len(df_original):
        print("⚠️ ADVERTENCIA: cambió el número de filas.")
        print(f"  Antes: {len(df_original)}")
        print(f"  Después: {len(df)}")
        print("  Revisa antes de subir a Excel Web.")
        print()

    if dry_run:
        print("DRY-RUN activo: no se modificó el Excel.")
        print("=" * 80)
        return

    print(f"Creando backup local: {backup_path}")
    shutil.copy2(excel_path, backup_path)

    print("Guardando Excel corregido...")
    safe_save_pandas(
        df,
        str(excel_path),
        sheet_name="Facturas",
        header=True,
        index=False,
    )

    print("Aplicando formato visual...")
    _aplicar_formato_visual_facturas()

    # Verificación final leyendo de nuevo el archivo guardado.
    df_final = pd.read_excel(excel_path, sheet_name="Facturas", engine="openpyxl")
    df_final_clean = _limpiar_dataframe_a_formato_largo(df_final)

    print()
    print("VERIFICACIÓN FINAL GUARDADA")
    print(f"  Filas: {len(df_final)}")
    print(f"  Facturas/grupos: {_contar_facturas(df_final_clean)}")
    print(f"  Celdas con texto tipo nan/null/none: {_contar_nan_texto(df_final)}")
    _print_counter("  Estado_calidad por fila:", _estado_counts(df_final))
    _print_counter("  Estado_calidad por factura:", _conceptos_por_estado(df_final_clean))

    print()
    print("✅ Corrección local completada.")
    print(f"Backup creado: {backup_path}")
    print("=" * 80)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Corrige facturas.xlsx limpiando nan y recalculando Estado_calidad."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Solo muestra el resumen antes/después. No guarda cambios.",
    )
    args = parser.parse_args()

    corregir_excel(dry_run=args.dry_run)


if __name__ == "__main__":
    main()
