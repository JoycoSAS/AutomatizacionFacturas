# services/m365/mail_graph.py
import os
import base64
import requests
from urllib.parse import quote
from dotenv import load_dotenv

from .token import get_access_token
from config import STORE_NAME  # fallback al UPN que ya tienes en config.py

load_dotenv()

GRAPH = "https://graph.microsoft.com/v1.0"

# --- Buzón: prioridad .env; si no, STORE_NAME ---
MAILBOX = (
    os.getenv("GRAPH_USER")
    or os.getenv("GRAPH_MAILBOX")
    or os.getenv("MAILBOX_UPN")
    or STORE_NAME
)
if not MAILBOX or not isinstance(MAILBOX, str):
    raise RuntimeError(
        "No hay buzón. Define GRAPH_USER/GRAPH_MAILBOX/MAILBOX_UPN en .env "
        "o usa STORE_NAME en config.py"
    )

TIMEOUT  = (15, 60)
_SESSION = requests.Session()


# ----------------------------
# Helpers HTTP / Autenticación
# ----------------------------
def _h():
    return {"Authorization": f"Bearer {get_access_token()}"}

def _user_segment() -> str:
    """Siempre usar /users/{mailbox}; evita /me. Codificamos solo el UPN."""
    return f"users/{quote(MAILBOX)}"

def _get(url, **kwargs):
    """GET con diagnóstico amigable si falla."""
    r = _SESSION.get(url, headers=_h(), timeout=kwargs.pop("timeout", TIMEOUT))
    if r.status_code >= 400:
        try:
            print("[Graph ERROR]", r.status_code, url, "->", r.json())
        except Exception:
            print("[Graph ERROR]", r.status_code, url, "->", r.text)
        r.raise_for_status()
    return r


# ------------------------------------
# BLOQUE: ZIPs desde Inbox (flujo base)
# ------------------------------------
def _categorias_ok(msg, required_categories=None):
    if not required_categories:
        return True
    cats = set([c.lower() for c in (msg.get("categories") or [])])
    return all(c.lower() in cats for c in required_categories)

def _listar_mensajes(max_messages=200, since_days=None):
    base = f"{GRAPH}/{_user_segment()}/messages"
    params = {
        "$select": "id,subject,hasAttachments,categories,receivedDateTime",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(max_messages, 500))
    }
    url = base + "?" + "&".join([f"{k}={quote(v)}" for k, v in params.items()])
    r = _get(url)
    return r.json().get("value", [])

def _listar_adjuntos(msg_id: str):
    """Lista adjuntos (todos los campos por defecto). NO usa $select para evitar 400."""
    mid = quote(msg_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/messages/{mid}/attachments"
    r = _get(url, timeout=(15, 120))
    return r.json().get("value", [])

def _listar_adjuntos_zip(msg_id: str):
    items = _listar_adjuntos(msg_id)
    out = []
    for a in items:
        name = (a.get("name") or "")
        cty  = (a.get("contentType") or "").lower()
        if name.lower().endswith(".zip") or "zip" in cty:
            out.append(a)
    return out

def _descargar_adjunto(msg_id: str, att_id: str, dest_path: str):
    """Descarga el adjunto (lee contentBytes desde GET /attachments/{id})."""
    mid = quote(msg_id, safe="")
    aid = quote(att_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/messages/{mid}/attachments/{aid}"
    r = _get(url, timeout=(15, 120))
    j = r.json()
    content = j.get("contentBytes")
    if not content:
        return False
    data = base64.b64decode(content)
    with open(dest_path, "wb") as f:
        f.write(data)
    return True

def descargar_zips_validos(temp_check_dir, destino_dir, read_all=False, max_messages=200,
                           since_days=None, required_categories=None):
    import zipfile, os
    os.makedirs(temp_check_dir, exist_ok=True)
    os.makedirs(destino_dir, exist_ok=True)

    msgs = _listar_mensajes(max_messages=max_messages, since_days=since_days)
    descargados = []

    for msg in msgs:
        if not msg.get("hasAttachments"):
            continue
        if not _categorias_ok(msg, required_categories):
            continue

        msg_id = msg["id"]
        atts = _listar_adjuntos_zip(msg_id)
        if not atts:
            continue

        for att in atts:
            name = att.get("name") or f"{att['id']}.zip"
            tmp_path = os.path.join(temp_check_dir, name)
            if not _descargar_adjunto(msg_id, att["id"], tmp_path):
                continue

            # ¿El ZIP trae XML?
            try:
                with zipfile.ZipFile(tmp_path, "r") as zf:
                    tiene_xml = any(m.filename.lower().endswith(".xml") for m in zf.infolist())
            except Exception:
                tiene_xml = False

            if not tiene_xml:
                try: os.remove(tmp_path)
                except Exception: pass
                continue

            dest_path = os.path.join(destino_dir, name)
            if not os.path.exists(dest_path):
                os.replace(tmp_path, dest_path)
                descargados.append(name)
            else:
                try: os.remove(tmp_path)
                except Exception: pass

    return descargados


# ------------------------------------------------
# BLOQUE: PDFs desde carpeta de “Facturas aprobadas”
# ------------------------------------------------
def get_folder_id_by_name(root_display: str, name: str) -> str | None:
    """Busca 'name' como hija directa de Inbox."""
    inbox_url = f"{GRAPH}/{_user_segment()}/mailFolders/inbox"
    r = _get(inbox_url)
    root_id = r.json()["id"]

    childs_url = f"{GRAPH}/{_user_segment()}/mailFolders/{root_id}/childFolders?$top=500"
    r = _get(childs_url)
    for item in r.json().get("value", []):
        if (item.get("displayName") or "").strip().lower() == name.strip().lower():
            return item["id"]
    return None

def find_folder_id_anywhere(name: str) -> str | None:
    """Búsqueda global por displayName en todo el buzón."""
    url = f"{GRAPH}/{_user_segment()}/mailFolders?$top=1000&$select=id,displayName"
    r = _get(url)
    for f in r.json().get("value", []):
        if (f.get("displayName") or "").strip().lower() == name.strip().lower():
            return f["id"]
    return None

def listar_mensajes_en_carpeta(folder_id: str, top: int = 200):
    fid = quote(folder_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/mailFolders/{fid}/messages"
    params = {
        "$select": "id,subject,hasAttachments,receivedDateTime,conversationId",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(top, 500))
    }
    q = "&".join([f"{k}={quote(v)}" for k, v in params.items()])
    r = _get(f"{url}?{q}")
    return r.json().get("value", [])

def listar_adjuntos_pdf(msg_id: str):
    """Lista SOLO los PDFs; descarga de archivo se hace con _descargar_adjunto()."""
    items = _listar_adjuntos(msg_id)
    pdfs = []
    for a in items:
        name = (a.get("name") or "").lower()
        cty  = (a.get("contentType") or "").lower()
        if name.endswith(".pdf") or "pdf" in cty:
            pdfs.append(a)
    return pdfs

def guardar_adjunto_base64(att_json: dict, dest_path: str) -> bool:
    """
    Mantenida por compatibilidad: intenta con contentBytes y, si no viene,
    retorna False para que el llamador use _descargar_adjunto().
    """
    content = att_json.get("contentBytes")
    if not content:
        return False
    data = base64.b64decode(content)
    with open(dest_path, "wb") as f:
        f.write(data)
    return True

def listar_mensajes_zip_inbox(top: int = 300):
    url = f"{GRAPH}/{_user_segment()}/messages"
    params = {
        "$select": "id,subject,hasAttachments,receivedDateTime",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(top, 500))
    }
    r = _get(f"{url}?{'&'.join([f'{k}={quote(v)}' for k,v in params.items()])}")
    return [m for m in r.json().get("value", []) if m.get("hasAttachments")]

def listar_adjuntos_zip(msg_id: str):
    """Lista SOLO los ZIPs; descarga con _descargar_adjunto()."""
    return _listar_adjuntos_zip(msg_id)

# API pública para descargar por id (útil desde controladores)
def descargar_adjunto_por_id(msg_id: str, att_id: str, dest_path: str) -> bool:
    return _descargar_adjunto(msg_id, att_id, dest_path)
