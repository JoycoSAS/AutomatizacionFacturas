import os
import sys
import time
import json
import hashlib
import datetime
import tempfile
from pathlib import Path
from urllib.parse import quote
from typing import Any, Iterable, Tuple

import requests
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import config  # noqa: F401  # carga .env del proyecto
except Exception:
    pass

from services.m365.token import get_access_token
from services.m365.sp_graph import SP_FOLDER as BASE_SP

VERSION_UPLOAD = "2026-06-12-UPLOAD-CIERRE-DIARIO-SP-V5-RECURSIVO-VERIFICACION-DATOS"
GRAPH = "https://graph.microsoft.com/v1.0"

DATA_DIR = ROOT / "data"
CIERRES_DIR = DATA_DIR / "cierres_diarios"
TMP_VERIFY_DIR = DATA_DIR / "_tmp_verificacion_sharepoint"

NOW = datetime.datetime.now()
FECHA = NOW.strftime("%Y-%m-%d")
MES = NOW.strftime("%Y-%m")

CIERRE_DIA_DIR = CIERRES_DIR / FECHA

SP_DRIVE_ID = (os.getenv("SP_DRIVE_ID") or "").strip()
SP_CIERRE_DIR_PRINCIPAL = f"{BASE_SP}/Backups/01_Cierres_Diarios/{MES}/{FECHA}".strip("/")

SP_BACKUP2_HOSTNAME = (os.getenv("SP_BACKUP2_HOSTNAME") or "").strip()
SP_BACKUP2_SITE_PATH = (os.getenv("SP_BACKUP2_SITE_PATH") or "").strip()
SP_BACKUP2_DRIVE_ID = (os.getenv("SP_BACKUP2_DRIVE_ID") or "").strip()
SP_BACKUP2_FOLDER = (os.getenv("SP_BACKUP2_FOLDER") or "").strip().strip("/")
SP_CIERRE_DIR_SECUNDARIA = (
    f"{SP_BACKUP2_FOLDER}/01_Cierres_Diarios/{MES}/{FECHA}".strip("/")
    if SP_BACKUP2_FOLDER
    else ""
)

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


def graph_download_path_content(drive_id: str, remote_path: str) -> bytes:
    url = f"{GRAPH}/drives/{encode_drive_id(drive_id)}/root:/{encode_path(remote_path)}:/content"
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


def debe_subir(p: Path) -> bool:
    if not p.is_file():
        return False
    if p.name in EXCLUIR_NOMBRES:
        return False
    if p.suffix.lower() in EXCLUIR_EXT:
        return False
    return True


def listar_archivos_recursivos(base: Path):
    return [p for p in sorted(base.rglob("*")) if debe_subir(p)]


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
    """
    Calcula un hash semántico del contenido del Excel.
    No compara el ZIP interno ni metadatos que SharePoint puede reescribir.
    Compara nombres de hojas y celdas con valores/formulas.
    """
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
    p = TMP_VERIFY_DIR / f"download_{FECHA}_{seguro}"
    p.write_bytes(data)
    return p


def verificar_archivo_subido(
    *,
    local: Path,
    rel: str,
    drive_id: str,
    remote_path: str,
    item: dict,
) -> bool:
    """
    Regla de seguridad:
    - JSON/TXT/otros: SHA256 binario exacto contra descarga desde SharePoint.
    - Excel: descarga el archivo subido y compara datos internos hoja/celda.
      No acepta solo existencia o tamaño.
    """
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
            if verificar_archivo_subido(
                local=local,
                rel=rel,
                drive_id=drive_id,
                remote_path=remote_path,
                item=item,
            ):
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


def subir_archivos_destino(nombre_destino: str, drive_id: str, carpeta_sp: str, archivos: list[Path]) -> bool:
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
    for local in archivos:
        rel = local.relative_to(CIERRE_DIA_DIR).as_posix()
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


def main() -> int:
    print("=" * 100)
    print("SUBIDA CIERRE DIARIO A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Fecha: {FECHA}")
    print(f"Carpeta local: {CIERRE_DIA_DIR}")
    print(f"Ruta principal:   {SP_CIERRE_DIR_PRINCIPAL}")
    print(f"Ruta secundaria:  {SP_CIERRE_DIR_SECUNDARIA}")
    print("-" * 100)

    if not CIERRE_DIA_DIR.exists():
        print("❌ No existe carpeta de cierre diario local.")
        return 1

    archivos = listar_archivos_recursivos(CIERRE_DIA_DIR)
    if not archivos:
        print("❌ No hay archivos válidos para subir en el cierre diario local.")
        return 1

    print(f"📦 Archivos a subir recursivamente: {len(archivos)}")
    for p in archivos:
        print(f"   - {p.relative_to(CIERRE_DIA_DIR).as_posix()} ({p.stat().st_size} bytes)")

    if not SP_DRIVE_ID:
        print("❌ Falta SP_DRIVE_ID en .env para la ruta principal.")
        return 1
    if not SP_CIERRE_DIR_SECUNDARIA:
        print("❌ Falta SP_BACKUP2_FOLDER en .env para la ruta secundaria.")
        return 1

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
        SP_CIERRE_DIR_PRINCIPAL,
        archivos,
    )
    ok_secundaria = subir_archivos_destino(
        "SECUNDARIA_CONTROL_INTERNO",
        drive_secundario,
        SP_CIERRE_DIR_SECUNDARIA,
        archivos,
    )

    print("-" * 100)
    if ok_principal and ok_secundaria:
        print("✅ Subida de cierre diario a SharePoint terminada correctamente en ambas rutas.")
        print("✅ Verificación aplicada:")
        print("   - JSON/TXT/otros: SHA256 exacto después de descargar desde SharePoint.")
        print("   - Excel: comparación de datos internos hoja/celda después de descargar desde SharePoint.")
        print("=" * 100)
        return 0

    print("❌ Subida de cierre diario terminó con errores en una o más rutas.")
    print("⚠️ No borres ni archives localmente hasta revisar el error.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
