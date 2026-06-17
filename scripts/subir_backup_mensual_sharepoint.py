import os
import sys
import time
import json
import hashlib
import datetime
from pathlib import Path
from urllib.parse import quote
from typing import Optional, Any

import requests

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import config  # noqa: F401  # Carga .env del proyecto
except Exception:
    pass

from services.m365.token import get_access_token
from services.m365.sp_graph import SP_FOLDER as BASE_SP

VERSION_UPLOAD = "2026-06-17-UPLOAD-BACKUP-MENSUAL-SP-V4-HASH-EXACTO"
GRAPH = "https://graph.microsoft.com/v1.0"

DATA_DIR = ROOT / "data"
BACKUPS_DIR = DATA_DIR / "backups_mensuales"
TMP_VERIFY_DIR = DATA_DIR / "_tmp_verificacion_sharepoint_mensual"

NOW = datetime.datetime.now()
MES = NOW.strftime("%Y-%m")
MES_FILE = NOW.strftime("%Y_%m")
MES_DIR = BACKUPS_DIR / MES

SP_DRIVE_ID = (os.getenv("SP_DRIVE_ID") or "").strip()
SP_BACKUP_ROOT = (os.getenv("SP_BACKUP_ROOT") or f"{BASE_SP}/Backups").strip("/")
SP_BACKUP_MENSUALES_DIR = (os.getenv("SP_BACKUP_MENSUALES_DIR") or "02_Backups_Mensuales").strip("/")
SP_MENSUAL_DIR_PRINCIPAL = f"{SP_BACKUP_ROOT}/{SP_BACKUP_MENSUALES_DIR}/{MES}".strip("/")

SP_BACKUP2_HOSTNAME = (os.getenv("SP_BACKUP2_HOSTNAME") or "").strip()
SP_BACKUP2_SITE_PATH = (os.getenv("SP_BACKUP2_SITE_PATH") or "").strip()
SP_BACKUP2_DRIVE_ID = (os.getenv("SP_BACKUP2_DRIVE_ID") or "").strip()
SP_BACKUP2_FOLDER = (os.getenv("SP_BACKUP2_FOLDER") or "").strip().strip("/")
SP_BACKUP2_MENSUALES_DIR = (os.getenv("SP_BACKUP2_MENSUALES_DIR") or "02_Backups_Mensuales").strip("/")
SP_MENSUAL_DIR_SECUNDARIA = (
    f"{SP_BACKUP2_FOLDER}/{SP_BACKUP2_MENSUALES_DIR}/{MES}".strip("/")
    if SP_BACKUP2_FOLDER
    else ""
)

EXCLUIR_NOMBRES = {".env"}
EXCLUIR_EXT = {".tmp", ".lock"}


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


def buscar_ultimo_zip_mensual() -> Optional[Path]:
    if not MES_DIR.exists():
        return None
    archivos = sorted(
        MES_DIR.glob(f"backup_mensual_{MES_FILE}_*.zip"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return archivos[0] if archivos else None


def extraer_timestamp_backup(zip_file: Path) -> str:
    # backup_mensual_2026_06_20260617_110935.zip -> 20260617_110935
    stem = zip_file.stem
    prefix = f"backup_mensual_{MES_FILE}_"
    if stem.startswith(prefix):
        return stem[len(prefix):]
    return ""


def archivos_backup_actual() -> list[Path]:
    zip_file = buscar_ultimo_zip_mensual()
    if not zip_file:
        raise RuntimeError(f"No se encontró ZIP mensual en {MES_DIR} con patrón backup_mensual_{MES_FILE}_*.zip")

    ts = extraer_timestamp_backup(zip_file)
    archivos = [zip_file]

    manifest = MES_DIR / f"manifest_backup_mensual_{MES_FILE}_{ts}.json" if ts else None
    resumen = MES_DIR / f"RESUMEN_BACKUP_MENSUAL_{MES}_{ts}.txt" if ts else None

    if manifest and manifest.exists():
        archivos.append(manifest)
    else:
        raise RuntimeError(f"No se encontró manifest correspondiente al ZIP: {manifest}")

    if resumen and resumen.exists():
        archivos.append(resumen)
    else:
        print(f"⚠️ No se encontró resumen mensual correspondiente: {resumen}")
        print("   Se continuará con ZIP + manifest.")

    return archivos


def verificar_archivo_subido(local: Path, drive_id: str, item: dict) -> bool:
    item_id = item.get("id")
    if not item_id:
        raise RuntimeError("Graph no devolvió item.id; no se puede verificar descarga exacta.")

    remote_bytes = graph_download_item_content(drive_id, item_id)
    hash_local = sha256_file(local)
    hash_sp = sha256_bytes(remote_bytes)

    if hash_local != hash_sp:
        print(f"❌ Hash distinto: {local.name}")
        print(f"   SHA256 local: {hash_local}")
        print(f"   SHA256 SP:    {hash_sp}")
        print(f"   bytes_local={local.stat().st_size} | bytes_sp={len(remote_bytes)}")
        return False

    print(f"✅ Archivo verificado por SHA256 exacto: {local.name} ({local.stat().st_size} bytes)")
    return True


def subir_y_verificar_con_reintentos(nombre_destino: str, drive_id: str, remote_path: str, local: Path) -> bool:
    item = graph_put_content(drive_id, remote_path, local)

    intentos = 4
    espera = 2
    ultimo_error = None
    for intento in range(1, intentos + 1):
        try:
            if verificar_archivo_subido(local, drive_id, item):
                return True
        except Exception as e:
            ultimo_error = e
            print(f"⚠️ Verificación intento {intento}/{intentos} falló para {local.name}: {e}")

        if intento < intentos:
            print(f"   Reintentando verificación en {espera}s...")
            time.sleep(espera)
            espera *= 2

    if ultimo_error:
        print(f"❌ Verificación definitiva fallida para {local.name}: {ultimo_error}")
    else:
        print(f"❌ Verificación definitiva fallida para {local.name}.")
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
        remote_path = f"{carpeta_sp}/{local.name}".strip("/")
        try:
            print("☁️ Subiendo a SharePoint:")
            print(f"   Destino: {nombre_destino}")
            print(f"   Local:   {local}")
            print(f"   SP:      {remote_path}")
            ok = subir_y_verificar_con_reintentos(nombre_destino, drive_id, remote_path, local)
            if not ok:
                ok_todos = False
        except Exception as e:
            print(f"❌ Error subiendo/verificando {local.name} a {nombre_destino}: {e}")
            ok_todos = False

    return ok_todos


def main() -> int:
    print("=" * 100)
    print("SUBIDA BACKUP MENSUAL A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Mes: {MES}")
    print(f"Carpeta local: {MES_DIR}")
    print(f"Ruta principal:   {SP_MENSUAL_DIR_PRINCIPAL}")
    print(f"Ruta secundaria:  {SP_MENSUAL_DIR_SECUNDARIA}")
    print("-" * 100)

    try:
        archivos = archivos_backup_actual()
    except Exception as e:
        print(f"❌ No se pudo preparar backup mensual para subida: {e}")
        return 1

    print(f"📦 Archivos mensuales a subir: {len(archivos)}")
    for p in archivos:
        print(f"   - {p.name} ({p.stat().st_size} bytes) | sha256={sha256_file(p)}")

    if not SP_DRIVE_ID:
        print("❌ Falta SP_DRIVE_ID en .env para la ruta principal.")
        return 1
    if not SP_MENSUAL_DIR_SECUNDARIA:
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
        SP_MENSUAL_DIR_PRINCIPAL,
        archivos,
    )
    ok_secundaria = subir_archivos_destino(
        "SECUNDARIA_CONTROL_INTERNO",
        drive_secundario,
        SP_MENSUAL_DIR_SECUNDARIA,
        archivos,
    )

    print("-" * 100)
    if ok_principal and ok_secundaria:
        print("✅ Subida mensual terminada correctamente en AMBAS rutas.")
        print("✅ Verificación aplicada: ZIP/JSON/TXT descargados desde SharePoint y comparados por SHA256 exacto.")
        print("=" * 100)
        return 0

    print("❌ Subida mensual terminó con errores en una o más rutas.")
    print("⚠️ No borres ni archives localmente hasta revisar el error.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
