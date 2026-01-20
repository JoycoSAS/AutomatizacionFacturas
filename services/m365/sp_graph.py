# services/m365/sp_graph.py
import os
import json
import time
import requests
from pathlib import Path
from urllib.parse import quote
from dotenv import load_dotenv

from .token import get_access_token

load_dotenv()

GRAPH = "https://graph.microsoft.com/v1.0"
DRIVE_ID = os.getenv("SP_DRIVE_ID")
SP_FOLDER = (os.getenv("SP_FOLDER") or "").strip().strip("/")
SSL_VERIFY = (os.getenv("SSL_VERIFY", "true").lower() == "true")
TIMEOUT = (15, 60)

_SESSION = requests.Session()


def sp_join(*parts: str) -> str:
    cleaned = []
    for p in parts:
        if not p:
            continue
        cleaned.append(str(p).replace("\\", "/").strip("/"))
    return "/".join([p for p in cleaned if p])


def _h(ct: str | None = None) -> dict:
    h = {"Authorization": f"Bearer {get_access_token()}"}
    if ct:
        h["Content-Type"] = ct
    return h


def _req(call, max_retries: int = 4):
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


def get_item_by_path(rel_path: str) -> dict:
    """
    Devuelve el DriveItem JSON (incluye id) de un archivo/carpeta dada su ruta relativa en el drive.
    Ej: "Innovacion/08. Pruebas proyectos/autoFacturas/excel/facturas.xlsx"
    """
    rel_path = rel_path.replace("\\", "/").strip("/")
    url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(rel_path)}"
    r = _req(lambda: _SESSION.get(url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY))
    return r.json()


def ensure_folder(rel_path: str):
    if not rel_path:
        return
    rel_path = rel_path.replace("\\", "/").strip("/")
    parts = [p for p in rel_path.split("/") if p]
    current = ""
    for seg in parts:
        current = f"{current}/{seg}" if current else seg

        get_url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(current)}"
        r = _SESSION.get(get_url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY)
        if r.status_code == 200:
            continue

        if r.status_code == 404:
            parent = "/".join(current.split("/")[:-1]).strip("/")
            post_url = (
                f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(parent)}:/children"
                if parent else f"{GRAPH}/drives/{DRIVE_ID}/root/children"
            )
            payload = {"name": seg, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
            _req(lambda: _SESSION.post(
                post_url, headers=_h("application/json"),
                data=json.dumps(payload), timeout=TIMEOUT, verify=SSL_VERIFY
            ))
            continue

        _req(lambda: _SESSION.get(get_url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY))


def _exists(rel_path: str) -> bool:
    rel_path = rel_path.replace("\\", "/").strip("/")
    url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(rel_path)}"
    r = _SESSION.get(url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY)
    return r.status_code == 200


def upload_small_file(local_path: str, dest_rel_path: str, mode: str = "replace"):
    dest_rel_path = dest_rel_path.replace("\\", "/").strip("/")
    ensure_folder(os.path.dirname(dest_rel_path))

    if mode == "skip" and _exists(dest_rel_path):
        print(f"   ⏭️  (skip) Ya existe en SP: {dest_rel_path}")
        return {"skipped": True, "name": os.path.basename(dest_rel_path)}

    put_url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(dest_rel_path)}:/content"
    with open(local_path, "rb") as f:
        data = f.read()

    r = _req(lambda: _SESSION.put(
        put_url, headers=_h(), data=data,
        timeout=(TIMEOUT[0], 300), verify=SSL_VERIFY
    ))
    try:
        return r.json()
    except Exception:
        return {"ok": True, "dest": dest_rel_path}


def upload_directory(local_dir: str, dest_rel_dir: str, mode: str = "replace"):
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


def download_small_file(sp_relative_path: str, local_path: str) -> bool:
    try:
        sp_path = sp_relative_path.strip().replace("\\", "/").strip("/")
        url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(sp_path)}:/content"
        r = _req(lambda: _SESSION.get(
            url, headers=_h(), timeout=(TIMEOUT[0], 300), verify=SSL_VERIFY
        ))
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(r.content)
        return True
    except Exception as e:
        print(f"[SP] Error descargando {sp_relative_path}: {e}")
        return False
