# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Reemplazo controlado del Excel activo de SharePoint por el facturas.xlsx limpio.

Modos:
- --dry-run:
  valida configuración, ruta destino y estructura local.
  No sube ni reemplaza nada.

- --upload-activo --confirmar REEMPLAZAR_EXCEL_ACTIVO_TRIMESTRAL:
  reemplaza el Excel activo de SharePoint únicamente si data/facturas.xlsx
  está limpio y tiene la estructura esperada.

Este script está pensado para ejecutarse DESPUÉS del cierre trimestral real local/VPS.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
import sys
from pathlib import Path
from typing import Any, Tuple
from urllib.parse import quote

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import config  # noqa: F401
except Exception:
    pass

from services.m365.token import get_access_token

try:
    from services.m365.sp_graph import SP_FOLDER as BASE_SP
except Exception:
    BASE_SP = ""

DATA_DIR = ROOT / "data"
FACTURAS_PATH = DATA_DIR / "facturas.xlsx"
TMP_VERIFY_DIR = DATA_DIR / "_tmp_verificacion_excel_activo_trimestral"

VERSION = "2026-06-18-REEMPLAZO-EXCEL-ACTIVO-TRIMESTRAL-SP-V1"
GRAPH = "https://graph.microsoft.com/v1.0"
CONFIRMACION = "REEMPLAZAR_EXCEL_ACTIVO_TRIMESTRAL"

load_dotenv(ROOT / ".env")

BASE_SP = (BASE_SP or os.getenv("SP_FOLDER") or "").strip().strip("/")
SP_DRIVE_ID = (os.getenv("SP_DRIVE_ID") or "").strip()

HEADERS_ESPERADOS = [
    "Radicado",
    "ProyectoProceso",
    "Archivo",
    "Empresa emisora",
    "CUFE",
    "Ciudad emisora",
    "Código ciudad",
    "NIT",
    "Cliente",
    "Número de factura",
    "Año",
    "Mes",
    "Día",
    "Tipo de contribuyente",
    "Actividad económica",
    "DESCRIPCIÓN",
    "Concepto",
    "VALOR",
    "Estado_calidad",
]


def ssl_verify() -> bool:
    return (os.getenv("SSL_VERIFY") or "true").strip().lower() not in {"0", "false", "no", "off"}


def headers_graph() -> dict:
    token = get_access_token()
    return {"Authorization": f"Bearer {token}"}


def encode_path(path: str) -> str:
    return quote(str(path).strip("/"), safe="/")


def encode_drive_id(drive_id: str) -> str:
    return quote(str(drive_id), safe="!")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def validar_config() -> None:
    if not SP_DRIVE_ID:
        raise RuntimeError("Falta SP_DRIVE_ID en .env.")
    if not BASE_SP:
        raise RuntimeError("Falta SP_FOLDER en .env.")
    if not FACTURAS_PATH.exists():
        raise RuntimeError(f"No existe el Excel local activo: {FACTURAS_PATH}")


def validar_excel_limpio(path: Path) -> dict:
    wb = load_workbook(path, read_only=False, data_only=False)
    try:
        if "Facturas" not in wb.sheetnames:
            raise RuntimeError("El Excel no tiene la hoja requerida: Facturas")

        ws = wb["Facturas"]

        headers = [cell.value for cell in ws[1]]
        if headers != HEADERS_ESPERADOS:
            raise RuntimeError(
                "Los encabezados del Excel no coinciden con la estructura esperada.\n"
                f"Detectados: {headers}\n"
                f"Esperados: {HEADERS_ESPERADOS}"
            )

        tablas = list(ws.tables.keys())
        tbl_ref = ws.tables["TblFacturas"].ref if "TblFacturas" in ws.tables else None

        info = {
            "archivo": str(path),
            "hojas": wb.sheetnames,
            "hoja_principal": ws.title,
            "filas": ws.max_row,
            "columnas": ws.max_column,
            "tablas": tablas,
            "tbl_facturas_ref": tbl_ref,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        }

        if ws.max_row != 1:
            raise RuntimeError(
                "El Excel local NO está limpio. Para reemplazar SharePoint debe tener solo encabezados.\n"
                f"Filas detectadas: {ws.max_row}. Se esperaba: 1.\n"
                "Esto es normal antes de ejecutar el cierre trimestral real local/VPS."
            )

        if ws.max_column != 19:
            raise RuntimeError(f"Columnas inválidas. Detectadas: {ws.max_column}. Esperadas: 19.")

        if tbl_ref != "A1:S1":
            raise RuntimeError(f"Tabla TblFacturas inválida. Detectada: {tbl_ref}. Esperada: A1:S1.")

        return info
    finally:
        wb.close()


def validar_excel_estructura_basica(path: Path) -> dict:
    """
    Valida estructura sin exigir que esté limpio.
    Sirve para el dry-run actual, cuando data/facturas.xlsx todavía tiene datos.
    """
    wb = load_workbook(path, read_only=False, data_only=False)
    try:
        if "Facturas" not in wb.sheetnames:
            raise RuntimeError("El Excel no tiene la hoja requerida: Facturas")

        ws = wb["Facturas"]
        headers = [cell.value for cell in ws[1]]
        tablas = list(ws.tables.keys())
        tbl_ref = ws.tables["TblFacturas"].ref if "TblFacturas" in ws.tables else None

        if headers != HEADERS_ESPERADOS:
            raise RuntimeError("Los encabezados del Excel activo no coinciden con la estructura esperada.")

        return {
            "archivo": str(path),
            "hojas": wb.sheetnames,
            "hoja_principal": ws.title,
            "filas": ws.max_row,
            "columnas": ws.max_column,
            "tablas": tablas,
            "tbl_facturas_ref": tbl_ref,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
            "esta_limpio_para_reemplazo": ws.max_row == 1 and ws.max_column == 19 and tbl_ref == "A1:S1",
        }
    finally:
        wb.close()


def normalizar_valor_excel(v: Any) -> str:
    if v is None:
        return "<NULL>"
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return v.isoformat()
    if isinstance(v, float):
        return repr(v)
    return f"{type(v).__name__}:{str(v)}"


def digest_datos_excel(path: Path) -> Tuple[str, dict]:
    h = hashlib.sha256()
    resumen = {
        "sheets": [],
        "non_empty_cells": 0,
    }

    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        sheet_names = list(wb.sheetnames)
        h.update(json.dumps(sheet_names, ensure_ascii=False).encode("utf-8"))

        for ws in wb.worksheets:
            sheet_info = {
                "title": ws.title,
                "max_row": int(ws.max_row or 0),
                "max_column": int(ws.max_column or 0),
                "non_empty_cells": 0,
            }

            h.update(f"\n[SHEET]{ws.title}|{ws.max_row}|{ws.max_column}".encode("utf-8"))

            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if v is None:
                        continue

                    sheet_info["non_empty_cells"] += 1
                    resumen["non_empty_cells"] += 1

                    payload = f"{ws.title}|{cell.coordinate}|{normalizar_valor_excel(v)}\n"
                    h.update(payload.encode("utf-8", errors="replace"))

            resumen["sheets"].append(sheet_info)
    finally:
        wb.close()

    return h.hexdigest(), resumen


def graph_put_content(drive_id: str, remote_path: str, local_file: Path) -> dict:
    url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/root:/{encode_path(remote_path)}:/content"
    data = local_file.read_bytes()

    r = requests.put(
        url,
        headers=headers_graph(),
        data=data,
        timeout=300,
        verify=ssl_verify(),
    )

    if r.status_code not in (200, 201):
        raise RuntimeError(f"PUT {r.status_code} {url} -> {r.text[:500]}")

    return r.json()


def graph_download_item_content(drive_id: str, item_id: str) -> bytes:
    url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/items/{quote(item_id, safe='')}/content"

    r = requests.get(
        url,
        headers=headers_graph(),
        timeout=300,
        verify=ssl_verify(),
        allow_redirects=True,
    )

    if r.status_code != 200:
        raise RuntimeError(f"DOWNLOAD {r.status_code} {url} -> {r.text[:500]}")

    return r.content


def escribir_temporal_descarga(data: bytes) -> Path:
    TMP_VERIFY_DIR.mkdir(parents=True, exist_ok=True)
    p = TMP_VERIFY_DIR / "download_excel_activo_trimestral_facturas.xlsx"
    p.write_bytes(data)
    return p


def verificar_excel_subido(local: Path, drive_id: str, item: dict) -> bool:
    item_id = item.get("id")
    if not item_id:
        raise RuntimeError("Graph no devolvió item.id; no se puede verificar descarga.")

    remote_bytes = graph_download_item_content(drive_id, item_id)
    tmp = escribir_temporal_descarga(remote_bytes)

    try:
        validar_excel_limpio(tmp)

        digest_local, resumen_local = digest_datos_excel(local)
        digest_sp, resumen_sp = digest_datos_excel(tmp)

        if digest_local != digest_sp:
            print("❌ El Excel activo subido tiene datos distintos al local.")
            print(f"Digest local: {digest_local}")
            print(f"Digest SP:    {digest_sp}")
            print(f"Resumen local: {json.dumps(resumen_local, ensure_ascii=False)}")
            print(f"Resumen SP:    {json.dumps(resumen_sp, ensure_ascii=False)}")
            return False

        print("✅ Excel activo de SharePoint verificado correctamente.")
        print(f"   Celdas no vacías: {resumen_local.get('non_empty_cells')}")
        print(f"   Hojas: {len(resumen_local.get('sheets', []))}")
        print(f"   Bytes local: {local.stat().st_size}")
        print(f"   Bytes SP: {len(remote_bytes)}")
        return True

    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


def remote_excel_activo_path() -> str:
    return f"{BASE_SP}/excel/facturas.xlsx".strip("/")


def imprimir_plan(info_basica: dict) -> None:
    destino = remote_excel_activo_path()

    print("✅ Configuración detectada.")
    print(f"SP_DRIVE_ID principal: {SP_DRIVE_ID or '(vacío)'}")
    print(f"SP_FOLDER principal: {BASE_SP or '(vacío)'}")
    print("-" * 100)
    print("Excel local activo detectado:")
    print(f"Archivo: {info_basica['archivo']}")
    print(f"Hojas: {info_basica['hojas']}")
    print(f"Hoja principal: {info_basica['hoja_principal']}")
    print(f"Filas: {info_basica['filas']}")
    print(f"Columnas: {info_basica['columnas']}")
    print(f"Tablas: {info_basica['tablas']}")
    print(f"TblFacturas ref: {info_basica['tbl_facturas_ref']}")
    print(f"Está limpio para reemplazo: {info_basica['esta_limpio_para_reemplazo']}")
    print("-" * 100)
    print("PLAN:")
    print("1. Validar que data\\facturas.xlsx esté limpio.")
    print("2. Reemplazar Excel activo en SharePoint principal:")
    print(f"   {destino}")
    print("3. Descargar desde SharePoint y verificar:")
    print("   - Hoja Facturas")
    print("   - 19 columnas")
    print("   - TblFacturas A1:S1")
    print("   - Digest de datos internos")


def ejecutar_upload_activo() -> int:
    validar_config()

    info_limpio = validar_excel_limpio(FACTURAS_PATH)
    destino = remote_excel_activo_path()

    print("✅ Excel local limpio validado para reemplazo.")
    print(f"Archivo: {FACTURAS_PATH}")
    print(f"Filas: {info_limpio['filas']}")
    print(f"Columnas: {info_limpio['columnas']}")
    print(f"TblFacturas ref: {info_limpio['tbl_facturas_ref']}")
    print("-" * 100)
    print("☁️ Reemplazando Excel activo de SharePoint...")
    print(f"SP: {destino}")

    item = graph_put_content(SP_DRIVE_ID, destino, FACTURAS_PATH)

    if verificar_excel_subido(FACTURAS_PATH, SP_DRIVE_ID, item):
        print("-" * 100)
        print("✅ Excel activo de SharePoint reemplazado y verificado correctamente.")
        print("=" * 100)
        return 0

    print("❌ No se pudo verificar correctamente el Excel activo reemplazado.")
    print("=" * 100)
    return 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Diagnostica sin reemplazar.")
    parser.add_argument("--upload-activo", action="store_true", help="Reemplaza Excel activo SharePoint.")
    parser.add_argument("--confirmar", default=None, help="Confirmación obligatoria para --upload-activo.")
    args = parser.parse_args()

    if args.dry_run and args.upload_activo:
        print("❌ Usa solo un modo: --dry-run o --upload-activo.")
        return 1

    modo = "UPLOAD_ACTIVO" if args.upload_activo else "DRY RUN"

    print("=" * 100)
    print(f"REEMPLAZO EXCEL ACTIVO TRIMESTRAL SHAREPOINT - {modo}")
    print("=" * 100)
    print(f"Versión: {VERSION}")
    print(f"Root: {ROOT}")
    print("-" * 100)

    try:
        validar_config()
        info_basica = validar_excel_estructura_basica(FACTURAS_PATH)
        imprimir_plan(info_basica)

        if not args.upload_activo:
            print("-" * 100)
            print("✅ DRY RUN finalizado. No se reemplazó ningún archivo.")
            print("=" * 100)
            return 0

        if args.confirmar != CONFIRMACION:
            print("-" * 100)
            print("❌ Reemplazo real bloqueado por falta de confirmación.")
            print("Para reemplazar el Excel activo usa:")
            print(f"python scripts\\reemplazar_excel_activo_trimestral_sharepoint.py --upload-activo --confirmar {CONFIRMACION}")
            print("=" * 100)
            return 1

        return ejecutar_upload_activo()

    except Exception as exc:
        print(f"❌ Error en reemplazo de Excel activo trimestral: {exc}")
        print("No se reemplazó ningún archivo en SharePoint.")
        print("=" * 100)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
