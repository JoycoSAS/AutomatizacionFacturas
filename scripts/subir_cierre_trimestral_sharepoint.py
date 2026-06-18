# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Subida de cierre trimestral a SharePoint.

Modos:
- --dry-run: solo diagnostica y calcula rutas. No sube archivos.
- --upload-cierre --confirmar SUBIR_CIERRE_TRIMESTRAL:
  sube el cierre trimestral local a SharePoint principal y secundario,
  con verificación fuerte.

Esta versión NO reemplaza el Excel activo de SharePoint.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
import sys
import time
from pathlib import Path
from typing import Any, Optional, Tuple
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
CIERRES_TRIMESTRALES_DIR = DATA_DIR / "cierres_trimestrales"
TMP_VERIFY_DIR = DATA_DIR / "_tmp_verificacion_sharepoint_trimestral"

VERSION_UPLOAD = "2026-06-18-UPLOAD-CIERRE-TRIMESTRAL-SP-V2-UPLOAD-CIERRE"
GRAPH = "https://graph.microsoft.com/v1.0"
CONFIRMACION_UPLOAD = "SUBIR_CIERRE_TRIMESTRAL"

load_dotenv(ROOT / ".env")

BASE_SP = (BASE_SP or os.getenv("SP_FOLDER") or "").strip().strip("/")

SP_DRIVE_ID = (os.getenv("SP_DRIVE_ID") or "").strip()
SP_BACKUP2_HOSTNAME = (os.getenv("SP_BACKUP2_HOSTNAME") or "").strip()
SP_BACKUP2_SITE_PATH = (os.getenv("SP_BACKUP2_SITE_PATH") or "").strip()
SP_BACKUP2_DRIVE_ID = (os.getenv("SP_BACKUP2_DRIVE_ID") or "").strip()
SP_BACKUP2_FOLDER = (os.getenv("SP_BACKUP2_FOLDER") or "").strip().strip("/")

EXCLUIR_NOMBRES = {".env"}
EXCLUIR_EXT = {".tmp", ".lock"}
EXCEL_EXT = {".xlsx", ".xlsm"}


def ssl_verify() -> bool:
    return (os.getenv("SSL_VERIFY") or "true").strip().lower() not in {"0", "false", "no", "off"}


def headers() -> dict:
    token = get_access_token()
    return {"Authorization": f"Bearer {token}"}


def h_json() -> dict:
    h = headers()
    h["Content-Type"] = "application/json"
    return h


def encode_path(path: str) -> str:
    return quote(str(path).strip("/"), safe="/")


def encode_drive_id(drive_id: str) -> str:
    return quote(str(drive_id), safe="!")


def graph_get(url: str, *, ok=(200,), timeout=60):
    r = requests.get(url, headers=headers(), timeout=timeout, verify=ssl_verify())
    if r.status_code not in ok:
        raise RuntimeError(f"GET {r.status_code} {url} -> {r.text[:500]}")
    return r


def graph_post(url: str, body: dict, *, ok=(200, 201), timeout=60):
    r = requests.post(url, headers=h_json(), json=body, timeout=timeout, verify=ssl_verify())
    if r.status_code not in ok:
        raise RuntimeError(f"POST {r.status_code} {url} -> {r.text[:500]}")
    return r


def graph_put_content(drive_id: str, remote_path: str, local_file: Path) -> dict:
    url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/root:/{encode_path(remote_path)}:/content"
    data = local_file.read_bytes()
    r = requests.put(url, headers=headers(), data=data, timeout=300, verify=ssl_verify())
    if r.status_code not in (200, 201):
        raise RuntimeError(f"PUT {r.status_code} {url} -> {r.text[:500]}")
    return r.json()


def graph_download_item_content(drive_id: str, item_id: str) -> bytes:
    url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/items/{quote(item_id, safe='')}/content"
    r = requests.get(url, headers=headers(), timeout=300, verify=ssl_verify(), allow_redirects=True)
    if r.status_code != 200:
        raise RuntimeError(f"DOWNLOAD {r.status_code} {url} -> {r.text[:500]}")
    return r.content


def existe_path(drive_id: str, remote_path: str) -> bool:
    if not remote_path.strip("/"):
        return True

    url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/root:/{encode_path(remote_path)}:"
    r = requests.get(url, headers=headers(), timeout=60, verify=ssl_verify())

    if r.status_code == 200:
        return True
    if r.status_code == 404:
        return False

    raise RuntimeError(f"GET {r.status_code} {url} -> {r.text[:500]}")


def crear_folder(drive_id: str, parent_path: str, folder_name: str) -> None:
    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail",
    }

    if parent_path.strip("/"):
        url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/root:/{encode_path(parent_path)}:/children"
    else:
        url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/root/children"

    r = requests.post(url, headers=h_json(), json=body, timeout=60, verify=ssl_verify())

    if r.status_code in (200, 201, 409):
        return

    raise RuntimeError(f"POST {r.status_code} {url} -> {r.text[:500]}")


def ensure_folder_recursive(drive_id: str, folder_path: str) -> None:
    folder_path = folder_path.strip("/")
    if not folder_path:
        return

    actual = ""
    for parte in [p for p in folder_path.split("/") if p]:
        siguiente = f"{actual}/{parte}".strip("/")
        if not existe_path(drive_id, siguiente):
            crear_folder(drive_id, actual, parte)
        actual = siguiente


def validar_drive_id(drive_id: str):
    if not drive_id:
        return None

    url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}?$select=id,name,webUrl"
    r = requests.get(url, headers=headers(), timeout=60, verify=ssl_verify())

    if r.status_code == 200:
        return r.json()

    print(f"⚠️ Drive ID no válido o no accesible: {drive_id}")
    print(f"   Respuesta Graph: {r.status_code} {r.text[:250]}")
    return None


def listar_drives_site(hostname: str, site_path: str):
    if not hostname or not site_path:
        return []

    site_path = site_path if site_path.startswith("/") else f"/{site_path}"
    url = f"{GRAPH}/sites/{hostname}:{site_path}:/drives?$select=id,name,webUrl"
    r = graph_get(url, timeout=60)
    return r.json().get("value", [])


def resolver_drive_secundario() -> str:
    drive = validar_drive_id(SP_BACKUP2_DRIVE_ID)
    if drive:
        print(f"✅ Drive secundario validado por ID: {drive.get('name')} | {drive.get('id')}")
        return drive["id"]

    print("🔎 Buscando drive secundario desde hostname/site_path...")
    drives = listar_drives_site(SP_BACKUP2_HOSTNAME, SP_BACKUP2_SITE_PATH)

    if not drives:
        raise RuntimeError("No se encontraron drives para el site secundario.")

    print("📚 Drives encontrados en site secundario:")
    for d in drives:
        print(f"   - {d.get('name')} | {d.get('id')}")

    preferidos = {"documentos", "documents", "shared documents"}
    elegido = None

    for d in drives:
        if (d.get("name") or "").strip().lower() in preferidos:
            elegido = d
            break

    if not elegido:
        elegido = drives[0]

    print(f"✅ Drive secundario resuelto: {elegido.get('name')} | {elegido.get('id')}")
    print("💡 Si este ID funciona, actualiza SP_BACKUP2_DRIVE_ID en .env con este valor.")
    return elegido["id"]


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


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
        "max_rows_total": 0,
        "max_cols_total": 0,
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

            resumen["max_rows_total"] += sheet_info["max_row"]
            resumen["max_cols_total"] += sheet_info["max_column"]

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


def escribir_temporal_descarga(rel: str, data: bytes) -> Path:
    TMP_VERIFY_DIR.mkdir(parents=True, exist_ok=True)
    seguro = rel.replace("/", "__").replace("\\", "__")
    p = TMP_VERIFY_DIR / f"download_trimestral_{seguro}"
    p.write_bytes(data)
    return p


def verificar_archivo_subido(
    *,
    local: Path,
    rel: str,
    drive_id: str,
    item: dict,
) -> bool:
    item_id = item.get("id")

    if not item_id:
        raise RuntimeError("Graph no devolvió item.id; no se puede verificar descarga exacta.")

    remote_bytes = graph_download_item_content(drive_id, item_id)

    if local.suffix.lower() in EXCEL_EXT:
        tmp = escribir_temporal_descarga(rel, remote_bytes)
        try:
            digest_local, resumen_local = digest_datos_excel(local)
            digest_sp, resumen_sp = digest_datos_excel(tmp)

            if digest_local != digest_sp:
                print(f"❌ Excel con datos distintos: {rel}")
                print(f"   Digest local: {digest_local}")
                print(f"   Digest SP:    {digest_sp}")
                print(f"   Resumen local: {json.dumps(resumen_local, ensure_ascii=False)}")
                print(f"   Resumen SP:    {json.dumps(resumen_sp, ensure_ascii=False)}")
                return False

            print(
                f"✅ Excel verificado por DATOS: {rel} | "
                f"celdas_no_vacias={resumen_local.get('non_empty_cells')} | "
                f"hojas={len(resumen_local.get('sheets', []))} | "
                f"bytes_local={local.stat().st_size} | bytes_sp={len(remote_bytes)}"
            )
            return True
        finally:
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass

    hash_local = sha256_file(local)
    hash_sp = sha256_bytes(remote_bytes)

    if hash_local != hash_sp:
        print(f"❌ Hash distinto: {rel}")
        print(f"   SHA256 local: {hash_local}")
        print(f"   SHA256 SP:    {hash_sp}")
        print(f"   bytes_local={local.stat().st_size} | bytes_sp={len(remote_bytes)}")
        return False

    print(f"✅ Archivo verificado por SHA256 exacto: {rel} ({local.stat().st_size} bytes)")
    return True


def subir_y_verificar_con_reintentos(nombre_destino: str, drive_id: str, remote_path: str, local: Path, rel: str) -> bool:
    item = graph_put_content(drive_id, remote_path, local)

    intentos = 4
    espera = 2
    ultimo_error = None

    for intento in range(1, intentos + 1):
        try:
            if verificar_archivo_subido(local=local, rel=rel, drive_id=drive_id, item=item):
                return True
        except Exception as e:
            ultimo_error = e
            print(f"⚠️ Verificación intento {intento}/{intentos} falló para {rel}: {e}")

        if intento < intentos:
            print(f"   Reintentando verificación en {espera}s...")
            time.sleep(espera)
            espera *= 2

    if ultimo_error:
        print(f"❌ Verificación definitiva fallida para {rel}: {ultimo_error}")
    else:
        print(f"❌ Verificación definitiva fallida para {rel}.")

    return False


def debe_subir(p: Path) -> bool:
    if not p.is_file():
        return False
    if p.name in EXCLUIR_NOMBRES:
        return False
    if p.suffix.lower() in EXCLUIR_EXT:
        return False
    return True


def buscar_ultimo_cierre_trimestral() -> dict:
    if not CIERRES_TRIMESTRALES_DIR.exists():
        raise RuntimeError(f"No existe carpeta de cierres trimestrales: {CIERRES_TRIMESTRALES_DIR}")

    excels = sorted(
        CIERRES_TRIMESTRALES_DIR.rglob("01_Excel_Cierre/facturas_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    if not excels:
        raise RuntimeError("No se encontró ningún Excel cerrado trimestral en 01_Excel_Cierre.")

    excel_cierre = excels[0]
    carpeta_excel = excel_cierre.parent
    carpeta_periodo = carpeta_excel.parent
    carpeta_soportes = carpeta_periodo / "02_Soportes_Tecnicos"

    periodo = carpeta_periodo.name
    anio = carpeta_periodo.parent.name

    archivos = [p for p in sorted(carpeta_periodo.rglob("*")) if debe_subir(p)]

    manifests = sorted(
        carpeta_soportes.glob("manifest_cierre_trimestral*.json"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    resumenes = sorted(
        carpeta_soportes.glob("RESUMEN_CIERRE_TRIMESTRAL*.txt"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    return {
        "anio": anio,
        "periodo": periodo,
        "carpeta_periodo": carpeta_periodo,
        "carpeta_excel": carpeta_excel,
        "carpeta_soportes": carpeta_soportes,
        "excel_cierre": excel_cierre,
        "manifest": manifests[0] if manifests else None,
        "resumen": resumenes[0] if resumenes else None,
        "archivos": archivos,
    }


def calcular_rutas_sp(cierre: dict) -> dict:
    ruta_principal_cierre = (
        f"{BASE_SP}/Backups/03_Cierres_Trimestrales/"
        f"{cierre['anio']}/{cierre['periodo']}"
    ).strip("/")

    ruta_secundaria_cierre = (
        f"{SP_BACKUP2_FOLDER}/03_Cierres_Trimestrales/"
        f"{cierre['anio']}/{cierre['periodo']}"
    ).strip("/")

    ruta_excel_activo_principal = f"{BASE_SP}/excel/facturas.xlsx".strip("/")

    return {
        "principal_cierre": ruta_principal_cierre,
        "secundaria_cierre": ruta_secundaria_cierre,
        "excel_activo_principal": ruta_excel_activo_principal,
    }


def imprimir_diagnostico(cierre: dict, rutas: dict) -> None:
    print("✅ Cierre trimestral local detectado.")
    print(f"Año: {cierre['anio']}")
    print(f"Periodo: {cierre['periodo']}")
    print(f"Carpeta periodo: {cierre['carpeta_periodo']}")
    print(f"Carpeta Excel cierre: {cierre['carpeta_excel']}")
    print(f"Carpeta soportes: {cierre['carpeta_soportes']}")
    print(f"Excel cierre: {cierre['excel_cierre']} ({cierre['excel_cierre'].stat().st_size} bytes)")

    if cierre["manifest"]:
        print(f"Manifest más reciente: {cierre['manifest']} ({cierre['manifest'].stat().st_size} bytes)")
    else:
        print("⚠️ No se encontró manifest.")

    if cierre["resumen"]:
        print(f"Resumen más reciente: {cierre['resumen']} ({cierre['resumen'].stat().st_size} bytes)")
    else:
        print("⚠️ No se encontró resumen.")

    print("-" * 100)
    print(f"📦 Archivos a subir: {len(cierre['archivos'])}")
    for p in cierre["archivos"]:
        rel = p.relative_to(cierre["carpeta_periodo"]).as_posix()
        print(f"   - {rel} ({p.stat().st_size} bytes)")

    print("-" * 100)
    print("✅ Configuración SharePoint detectada.")
    print(f"SP_DRIVE_ID principal: {SP_DRIVE_ID or '(vacio)'}")
    print(f"SP_FOLDER principal: {BASE_SP or '(vacio)'}")
    print(f"SP_BACKUP2_DRIVE_ID secundario: {SP_BACKUP2_DRIVE_ID or '(vacio)'}")
    print(f"SP_BACKUP2_FOLDER secundario: {SP_BACKUP2_FOLDER or '(vacio)'}")

    print("-" * 100)
    print("PLAN SHAREPOINT:")
    print("1. Subir cierre trimestral a ruta principal:")
    print(f"   {rutas['principal_cierre']}")
    print("2. Subir cierre trimestral a ruta secundaria:")
    print(f"   {rutas['secundaria_cierre']}")
    print("3. NO reemplazar todavía el Excel activo principal:")
    print(f"   {rutas['excel_activo_principal']}")


def validar_config_sp() -> None:
    if not SP_DRIVE_ID:
        raise RuntimeError("Falta SP_DRIVE_ID en .env.")
    if not BASE_SP:
        raise RuntimeError("Falta SP_FOLDER en .env.")
    if not SP_BACKUP2_DRIVE_ID and not (SP_BACKUP2_HOSTNAME and SP_BACKUP2_SITE_PATH):
        raise RuntimeError("Falta SP_BACKUP2_DRIVE_ID o SP_BACKUP2_HOSTNAME/SP_BACKUP2_SITE_PATH en .env.")
    if not SP_BACKUP2_FOLDER:
        raise RuntimeError("Falta SP_BACKUP2_FOLDER en .env.")


def subir_archivos_destino(nombre_destino: str, drive_id: str, carpeta_sp: str, cierre: dict) -> bool:
    print("-" * 100)
    print(f"📁 Verificando/creando destino: {nombre_destino}")
    print(f"   Drive ID: {drive_id}")
    print(f"   SP_DIR:   {carpeta_sp}")

    try:
        ensure_folder_recursive(drive_id, carpeta_sp)
        print(f"✅ Carpeta SharePoint verificada/creada: {nombre_destino}")
    except Exception as e:
        print(f"❌ No se pudo verificar/crear carpeta SharePoint en {nombre_destino}: {e}")
        return False

    ok_todos = True

    for local in cierre["archivos"]:
        rel = local.relative_to(cierre["carpeta_periodo"]).as_posix()
        remote_path = f"{carpeta_sp}/{rel}".strip("/")
        remote_dir = "/".join(remote_path.split("/")[:-1])

        try:
            ensure_folder_recursive(drive_id, remote_dir)

            print("☁️ Subiendo a SharePoint:")
            print(f"   Destino: {nombre_destino}")
            print(f"   Local:   {local}")
            print(f"   SP:      {remote_path}")

            ok = subir_y_verificar_con_reintentos(nombre_destino, drive_id, remote_path, local, rel)
            if not ok:
                ok_todos = False

        except Exception as e:
            print(f"❌ Error subiendo/verificando {rel} a {nombre_destino}: {e}")
            ok_todos = False

    return ok_todos


def ejecutar_upload_cierre(cierre: dict, rutas: dict) -> int:
    validar_config_sp()

    try:
        print("🔐 Validando drive principal...")
        principal = validar_drive_id(SP_DRIVE_ID)
        if not principal:
            print("❌ SP_DRIVE_ID principal no es válido o no es accesible.")
            return 1

        print(f"✅ Drive principal validado: {principal.get('name')} | {principal.get('id')}")

        print("🔐 Validando/resolviendo drive secundario...")
        drive_secundario = resolver_drive_secundario()

    except Exception as e:
        print(f"❌ Error validando drives: {e}")
        return 1

    ok_principal = subir_archivos_destino(
        "PRINCIPAL_CONTABILIDAD",
        SP_DRIVE_ID,
        rutas["principal_cierre"],
        cierre,
    )

    ok_secundaria = subir_archivos_destino(
        "SECUNDARIA_CONTROL_INTERNO",
        drive_secundario,
        rutas["secundaria_cierre"],
        cierre,
    )

    print("-" * 100)

    if ok_principal and ok_secundaria:
        print("✅ Subida de cierre trimestral a SharePoint terminada correctamente en ambas rutas.")
        print("✅ Verificación aplicada:")
        print("   - JSON/TXT/otros: SHA256 exacto después de descargar desde SharePoint.")
        print("   - Excel: comparación de datos internos hoja/celda después de descargar desde SharePoint.")
        print("⚠️ No se reemplazó el Excel activo de SharePoint en esta versión.")
        print("=" * 100)
        return 0

    print("❌ Subida de cierre trimestral terminó con errores en una o más rutas.")
    print("⚠️ No ejecutes reemplazo de Excel activo hasta revisar el error.")
    print("=" * 100)
    return 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Solo diagnostica, no sube archivos.")
    parser.add_argument("--upload-cierre", action="store_true", help="Sube cierre trimestral a SharePoint.")
    parser.add_argument("--confirmar", default=None, help="Confirmación obligatoria para --upload-cierre.")
    args = parser.parse_args()

    if args.dry_run and args.upload_cierre:
        print("❌ Usa solo un modo: --dry-run o --upload-cierre.")
        return 1

    modo = "UPLOAD_CIERRE" if args.upload_cierre else "DRY RUN"

    print("=" * 100)
    print(f"SUBIDA CIERRE TRIMESTRAL SHAREPOINT - {modo}")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print("-" * 100)

    try:
        cierre = buscar_ultimo_cierre_trimestral()
        rutas = calcular_rutas_sp(cierre)
        imprimir_diagnostico(cierre, rutas)

        if not args.upload_cierre:
            print("-" * 100)
            print("✅ DRY RUN SharePoint finalizado. No se subió ningún archivo.")
            print("=" * 100)
            return 0

        if args.confirmar != CONFIRMACION_UPLOAD:
            print("-" * 100)
            print("❌ Subida real bloqueada por falta de confirmación.")
            print("Para subir el cierre trimestral usa:")
            print(f"python scripts\\subir_cierre_trimestral_sharepoint.py --upload-cierre --confirmar {CONFIRMACION_UPLOAD}")
            print("=" * 100)
            return 1

        return ejecutar_upload_cierre(cierre, rutas)

    except Exception as exc:
        print(f"❌ Error en subida trimestral SharePoint: {exc}")
        print("No se subió ningún archivo.")
        print("=" * 100)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
