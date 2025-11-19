# services/m365/sp_graph.py
# -------------------------------------------
# Cliente mínimo para subir/descargar archivos y
# gestionar carpetas en SharePoint/OneDrive (Graph).
# Mantiene lo que ya tenías y añade descarga.
# -------------------------------------------

import os
import json
import time
import requests
from pathlib import Path
from urllib.parse import quote
from dotenv import load_dotenv

# ✅ Helper de token dentro del paquete m365
from .token import get_access_token

load_dotenv()

GRAPH     = "https://graph.microsoft.com/v1.0"
DRIVE_ID  = os.getenv("SP_DRIVE_ID")
SP_FOLDER = os.getenv("SP_FOLDER") or ""  # ruta base dentro del drive
TIMEOUT   = (15, 60)  # (connect, read)

_SESSION = requests.Session()


# ----------------------------
# Helpers HTTP / Autenticación
# ----------------------------
def _h(ct: str | None = None) -> dict:
    """Headers con Bearer; opcional Content-Type."""
    h = {"Authorization": f"Bearer {get_access_token()}"}
    if ct:
        h["Content-Type"] = ct
    return h


def _req(call, max_retries: int = 4):
    """
    Ejecuta una request con reintentos exponenciales en
    429/5xx. Traza error con detalle si falla.
    """
    attempt = 0
    while True:
        r = call()
        if r.status_code < 400:
            return r

        if r.status_code in (429, 500, 502, 503, 504) and attempt < max_retries:
            attempt += 1
            wait = r.headers.get("Retry-After")
            try:
                wait = float(wait)
            except Exception:
                wait = min(2 ** attempt, 15)
            time.sleep(wait)
            continue

        try:
            body = r.json()
        except Exception:
            body = r.text
        print(f"[Graph ERROR] {r.status_code} {r.request.method} {r.url} -> {body}")
        r.raise_for_status()


# -----------------
# Carpetas remotas
# -----------------
def ensure_folder(rel_path: str):
    """
    Crea recursivamente cada segmento bajo el drive.
    rel_path: 'carpeta1/carpeta 2/carpeta3' (SIN barra inicial ni final)
    Es idempotente: si ya existe, sigue.
    """
    if not rel_path:
        return

    rel_path = rel_path.replace("\\", "/").strip("/")
    parts = [p for p in rel_path.split("/") if p]
    current = ""
    for seg in parts:
        current = f"{current}/{seg}" if current else seg

        get_url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(current)}"
        r = _SESSION.get(get_url, headers=_h(), timeout=TIMEOUT)

        if r.status_code == 200:
            continue  # existe → avanzar

        if r.status_code == 404:
            parent = "/".join(current.split("/")[:-1]).strip("/")
            if parent:
                post_url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(parent)}:/children"
            else:
                post_url = f"{GRAPH}/drives/{DRIVE_ID}/root/children"

            payload = {
                "name": seg,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename",
            }
            _req(lambda: _SESSION.post(
                post_url,
                headers=_h("application/json"),
                data=json.dumps(payload),
                timeout=TIMEOUT,
            ))
            continue

        _req(lambda: _SESSION.get(get_url, headers=_h(), timeout=TIMEOUT))


def _exists(rel_path: str) -> bool:
    """Verifica si existe un item (archivo/carpeta) por ruta relativa dentro del drive."""
    rel_path = rel_path.replace("\\", "/").strip("/")
    url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(rel_path)}"
    r = _SESSION.get(url, headers=_h(), timeout=TIMEOUT)
    return r.status_code == 200


# -------------
# Subida archivos
# -------------
def upload_small_file(local_path: str, dest_rel_path: str, mode: str = "replace"):
    """
    Sube un archivo (tamaño pequeño/mediano) con PUT único.
    - Crea la carpeta destino si no existe.
    - mode: "replace" (default) o "skip"
    """
    dest_rel_path = dest_rel_path.replace("\\", "/").strip("/")
    ensure_folder(os.path.dirname(dest_rel_path))

    if mode == "skip" and _exists(dest_rel_path):
        print(f"   ⏭️  (skip) Ya existe en SP: {dest_rel_path}")
        return {"skipped": True, "name": os.path.basename(dest_rel_path)}

    put_url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(dest_rel_path)}:/content"
    with open(local_path, "rb") as f:
        data = f.read()

    r = _req(lambda: _SESSION.put(
        put_url,
        headers=_h(),
        data=data,
        timeout=(TIMEOUT[0], 300),
    ))
    try:
        return r.json()
    except Exception:
        return {"ok": True, "dest": dest_rel_path}


def upload_directory(local_dir: str, dest_rel_dir: str, mode: str = "replace"):
    """
    ⬆️ SUBIDA RECURSIVA de todo el contenido de `local_dir` a `dest_rel_dir` en SharePoint.
    """
    local_dir = Path(local_dir)
    dest_rel_dir = dest_rel_dir.replace("\\", "/").strip("/")

    print(f"[DEBUG] Subiendo a: {dest_rel_dir!r} (mode={mode})")
    if not local_dir.exists():
        print(f"[WARN] Carpeta local no existe: {local_dir}")
        return

    ensure_folder(dest_rel_dir)

    for root, dirs, files in os.walk(local_dir):
        root_p = Path(root)
        rel = root_p.relative_to(local_dir)
        rel_sp = dest_rel_dir if str(rel) == "." else f"{dest_rel_dir}/{str(rel).replace('\\', '/')}"

        ensure_folder(rel_sp)

        for fname in files:
            local_path = root_p / fname
            server_rel_path = f"{rel_sp}/{fname}".replace("\\", "/")
            print(f"   ⬆️  {local_path.name} -> {server_rel_path}")
            upload_small_file(str(local_path), server_rel_path, mode=mode)


# -------------
# Descarga archivo
# -------------
def download_small_file(sp_relative_path: str, local_path: str) -> bool:
    """
    Descarga un archivo de SharePoint por ruta relativa (la misma que usas al subir).
    Ej: 'SOPORTES/Temporal Vehiculos/Prueba de Facturas Daniel/excel/Aprobaciones_Facturas.xlsx'
    """
    try:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        sp_path = sp_relative_path.strip().replace("\\", "/").strip("/")
        url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(sp_path)}:/content"
        r = _req(lambda: _SESSION.get(
            url,
            headers=_h(),
            timeout=(TIMEOUT[0], 300),
            verify=False,   # ⚠️ según tu entorno
        ))
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(r.content)
        return True
    except Exception as e:
        print(f"[SP] Error descargando {sp_relative_path}: {e}")
        return False
