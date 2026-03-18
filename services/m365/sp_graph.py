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

# =========================
# SharePoint 1 (default)
# =========================
DRIVE_ID_DEFAULT = (os.getenv("SP_DRIVE_ID") or "").strip()

# ✅ BACKWARD COMPAT (para imports antiguos: ExcelWorkbookGraph, etc.)
DRIVE_ID = DRIVE_ID_DEFAULT  # NO BORRAR

SP_FOLDER = (os.getenv("SP_FOLDER") or "").strip().strip("/")

SSL_VERIFY = (os.getenv("SSL_VERIFY", "true").lower() == "true")

# (connect_timeout, read_timeout)
TIMEOUT = (20, 120)
UPLOAD_TIMEOUT = (20, 300)

_SESSION = requests.Session()
_SESSION.headers.update({"Accept": "application/json"})


# -------------------------
# Helpers
# -------------------------
def _drive(drive_id: str | None) -> str:
    d = (drive_id or DRIVE_ID_DEFAULT or "").strip()
    if not d:
        raise RuntimeError(
            "No hay DRIVE_ID configurado. Revisa SP_DRIVE_ID (y/o SP_DRIVE_ID_RADICADOS)."
        )
    return d


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


def _sleep_backoff(attempt: int, max_wait: float = 20.0):
    wait = min(2 ** max(1, attempt), max_wait)
    time.sleep(wait)


def _extract_error_body(response: requests.Response):
    try:
        return response.json()
    except Exception:
        try:
            return response.text
        except Exception:
            return "<sin body>"


def _req(call, max_retries: int = 4):
    """
    Wrapper general con retry para:
    - 429
    - 500/502/503/504
    - ReadTimeout / ConnectionError
    """
    attempt = 0

    while True:
        try:
            r = call()

            if r.status_code < 400:
                return r

            if r.status_code in (429, 500, 502, 503, 504) and attempt < max_retries:
                attempt += 1
                retry_after = r.headers.get("Retry-After")
                try:
                    wait = float(retry_after)
                except Exception:
                    wait = min(2 ** attempt, 20)

                print(
                    f"[Graph RETRY] HTTP {r.status_code}. "
                    f"Reintento {attempt}/{max_retries} en {wait:.1f}s"
                )
                time.sleep(wait)
                continue

            body = _extract_error_body(r)
            print(f"[Graph ERROR] {r.status_code} {r.request.method} {r.url} -> {body}")
            r.raise_for_status()

        except (requests.exceptions.ReadTimeout, requests.exceptions.ConnectionError) as e:
            if attempt < max_retries:
                attempt += 1
                wait = min(2 ** attempt, 20)
                print(
                    f"[Graph RETRY] {type(e).__name__}. "
                    f"Reintento {attempt}/{max_retries} en {wait:.1f}s"
                )
                time.sleep(wait)
                continue
            raise

        except requests.exceptions.RequestException:
            raise


# -------------------------
# Core: obtener items
# -------------------------
def get_item_by_path(rel_path: str, drive_id: str | None = None) -> dict:
    """
    Devuelve el DriveItem JSON (incluye id) de un archivo/carpeta dada su ruta relativa en el drive.
    """
    d = _drive(drive_id)
    rel_path = (rel_path or "").replace("\\", "/").strip("/")

    url = f"{GRAPH}/drives/{d}/root:/{quote(rel_path)}:/"
    r = _req(lambda: _SESSION.get(url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY))
    return r.json()


def _exists(rel_path: str, drive_id: str | None = None) -> bool:
    """
    Verifica existencia sin levantar excepción por 404.
    Si hay timeout/error de red, devuelve False.
    """
    d = _drive(drive_id)
    rel_path = (rel_path or "").replace("\\", "/").strip("/")
    url = f"{GRAPH}/drives/{d}/root:/{quote(rel_path)}:/"

    try:
        r = _SESSION.get(url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY)
        return r.status_code == 200
    except Exception:
        return False


# -------------------------
# Carpetas
# -------------------------
def ensure_folder(rel_path: str, drive_id: str | None = None, max_retries: int = 3):
    """
    Asegura una carpeta (y sus padres) en el drive indicado.
    Versión robusta ante timeouts de Graph.
    """
    rel_path = (rel_path or "").replace("\\", "/").strip("/")
    if not rel_path:
        return

    d = _drive(drive_id)
    parts = [p for p in rel_path.split("/") if p]
    current = ""

    for seg in parts:
        current = f"{current}/{seg}" if current else seg
        attempt = 0

        while True:
            try:
                get_url = f"{GRAPH}/drives/{d}/root:/{quote(current)}:/"
                r = _SESSION.get(get_url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY)

                if r.status_code == 200:
                    break

                if r.status_code == 404:
                    parent = "/".join(current.split("/")[:-1]).strip("/")
                    post_url = (
                        f"{GRAPH}/drives/{d}/root:/{quote(parent)}:/children"
                        if parent
                        else f"{GRAPH}/drives/{d}/root/children"
                    )

                    payload = {
                        "name": seg,
                        "folder": {},
                        "@microsoft.graph.conflictBehavior": "rename",
                    }

                    _req(
                        lambda: _SESSION.post(
                            post_url,
                            headers=_h("application/json"),
                            data=json.dumps(payload),
                            timeout=TIMEOUT,
                            verify=SSL_VERIFY,
                        )
                    )
                    break

                # para cualquier otro error HTTP dejamos que _req lo trate
                _req(lambda: _SESSION.get(get_url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY))
                break

            except (requests.exceptions.ReadTimeout, requests.exceptions.ConnectionError) as e:
                attempt += 1
                if attempt > max_retries:
                    print(f"[SP ensure_folder] Error definitivo en '{current}': {e}")
                    raise
                wait = min(2 ** attempt, 20)
                print(
                    f"[SP ensure_folder] Timeout/conexión en '{current}'. "
                    f"Reintento {attempt}/{max_retries} en {wait:.1f}s"
                )
                time.sleep(wait)

            except requests.exceptions.RequestException as e:
                attempt += 1
                if attempt > max_retries:
                    print(f"[SP ensure_folder] Error HTTP definitivo en '{current}': {e}")
                    raise
                wait = min(2 ** attempt, 20)
                print(
                    f"[SP ensure_folder] Error HTTP en '{current}'. "
                    f"Reintento {attempt}/{max_retries} en {wait:.1f}s"
                )
                time.sleep(wait)


# -------------------------
# Upload / Download
# -------------------------
def upload_small_file(
    local_path: str,
    dest_rel_path: str,
    mode: str = "replace",
    drive_id: str | None = None,
):
    """
    Sube un archivo pequeño con PUT ...:/content
    mode:
      - "replace": siempre reemplaza
      - "skip": si existe, no sube
    """
    d = _drive(drive_id)
    dest_rel_path = str(dest_rel_path).replace("\\", "/").strip("/")

    if not os.path.exists(local_path):
        raise FileNotFoundError(f"No existe el archivo local: {local_path}")

    parent = os.path.dirname(dest_rel_path).replace("\\", "/").strip("/")
    if parent:
        ensure_folder(parent, drive_id=d)

    if mode == "skip" and _exists(dest_rel_path, drive_id=d):
        print(f"   ⏭️  (skip) Ya existe en SP: {dest_rel_path}")
        return {"skipped": True, "name": os.path.basename(dest_rel_path)}

    put_url = f"{GRAPH}/drives/{d}/root:/{quote(dest_rel_path)}:/content"
    with open(local_path, "rb") as f:
        data = f.read()

    r = _req(
        lambda: _SESSION.put(
            put_url,
            headers=_h(),
            data=data,
            timeout=UPLOAD_TIMEOUT,
            verify=SSL_VERIFY,
        )
    )

    try:
        return r.json()
    except Exception:
        return {"ok": True, "dest": dest_rel_path}


def upload_directory(
    local_dir: str,
    dest_rel_dir: str,
    mode: str = "replace",
    drive_id: str | None = None,
):
    """
    Sube una carpeta completa (recursivo) usando upload_small_file por cada archivo.
    """
    d = _drive(drive_id)
    local_dir = Path(local_dir)
    dest_rel_dir = str(dest_rel_dir).replace("\\", "/").strip("/")

    print(f"[DEBUG] Subiendo a: {dest_rel_dir!r} (mode={mode})")
    if not local_dir.exists():
        print(f"[WARN] Carpeta local no existe: {local_dir}")
        return

    if dest_rel_dir:
        ensure_folder(dest_rel_dir, drive_id=d)

    for root, _, files in os.walk(local_dir):
        root_p = Path(root)
        rel = root_p.relative_to(local_dir)
        rel_sp = dest_rel_dir if str(rel) == "." else sp_join(dest_rel_dir, str(rel))

        if rel_sp:
            ensure_folder(rel_sp, drive_id=d)

        for fname in files:
            local_path = root_p / fname
            server_rel_path = sp_join(rel_sp, fname)
            print(f"   ⬆️  {local_path.name} -> {server_rel_path}")
            upload_small_file(str(local_path), server_rel_path, mode=mode, drive_id=d)


def download_small_file(sp_relative_path: str, local_path: str, drive_id: str | None = None) -> bool:
    """
    Descarga un archivo por ruta relativa dentro del drive.
    """
    try:
        d = _drive(drive_id)
        sp_path = str(sp_relative_path).strip().replace("\\", "/").strip("/")
        url = f"{GRAPH}/drives/{d}/root:/{quote(sp_path)}:/content"

        r = _req(
            lambda: _SESSION.get(
                url,
                headers=_h(),
                timeout=UPLOAD_TIMEOUT,
                verify=SSL_VERIFY,
            )
        )

        parent = os.path.dirname(local_path)
        if parent:
            os.makedirs(parent, exist_ok=True)

        with open(local_path, "wb") as f:
            f.write(r.content)

        return True
    except Exception as e:
        print(f"[SP] Error descargando {sp_relative_path}: {e}")
        return False


# -------------------------
# Listar children (para tests)
# -------------------------
def list_children(rel_path: str = "", drive_id: str | None = None, top: int = 999) -> list[dict]:
    """
    Lista hijos (archivos/carpetas) en una ruta del drive.
    - Si rel_path está vacío -> lista el root del drive.
    - Soporta top y paginación.
    """
    d = _drive(drive_id)
    rel_path = (rel_path or "").replace("\\", "/").strip("/")

    if rel_path:
        url = f"{GRAPH}/drives/{d}/root:/{quote(rel_path)}:/children?$top={int(top)}"
    else:
        url = f"{GRAPH}/drives/{d}/root/children?$top={int(top)}"

    items: list[dict] = []
    while True:
        r = _req(lambda: _SESSION.get(url, headers=_h(), timeout=TIMEOUT, verify=SSL_VERIFY))
        data = r.json() or {}
        items.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        url = next_link

    return items