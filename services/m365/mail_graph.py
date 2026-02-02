# services/m365/mail_graph.py
import os
import base64
import time
import requests
from urllib.parse import quote
from dotenv import load_dotenv
from datetime import datetime, timedelta, timezone

from .token import get_access_token
from config import STORE_NAME  # fallback al UPN que ya tienes en config.py

load_dotenv()

GRAPH = "https://graph.microsoft.com/v1.0"

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

TIMEOUT = (15, 60)
_SESSION = requests.Session()
_SESSION.headers.update({"Accept": "application/json"})


def _h(content_type: str | None = None, extra: dict | None = None):
    """Headers con Bearer actual, opcional Content-Type."""
    h = {"Authorization": f"Bearer {get_access_token()}"}
    if content_type:
        h["Content-Type"] = content_type
    if extra:
        h.update(extra)
    return h


def _user_segment() -> str:
    """Siempre usar /users/{mailbox}; evita /me. Codificamos solo el UPN."""
    return f"users/{quote(MAILBOX)}"


def _get_json_with_retries(url: str, retries: int = 2, timeout=TIMEOUT, headers_extra: dict | None = None):
    """
    GET + parse JSON con reintentos si la respuesta viene truncada/no-JSON.
    Devuelve dict (JSON) o None si no fue posible.
    """
    delay = 1.5
    for attempt in range(retries + 1):
        try:
            r = _SESSION.get(url, headers=_h(extra=headers_extra), timeout=timeout)
            if not r.ok:
                print(f"[Graph] HTTP {r.status_code} en {url}")
                try:
                    err = r.json()
                    print("[Graph] Error detallado:", err)
                except Exception:
                    pass
            else:
                try:
                    return r.json()
                except Exception as e:
                    print(f"[Graph] Respuesta no-JSON: {e}")
                    snippet = (r.text[:200] + "...") if isinstance(r.text, str) and len(r.text) > 200 else r.text
                    if snippet:
                        print(f"→ Cuerpo parcial: {snippet}")
        except requests.RequestException as e:
            print(f"[Graph] Error de red/timeout al llamar a Graph: {e}")

        if attempt < retries:
            print(f"[Graph] Reintentando ({attempt + 1}/{retries}) en {delay:.1f}s…")
            time.sleep(delay)
            delay *= 2

    return None


def _categorias_ok(msg, required_categories=None):
    if not required_categories:
        return True
    cats = set([c.lower() for c in (msg.get("categories") or [])])
    return all(c.lower() in cats for c in required_categories)


# ----------------------------
# Listar mensajes (Inbox)
# ----------------------------
def _listar_mensajes(max_messages=200, since_days=None):
    """
    Lista mensajes del Inbox (ordenados desc).
    ✅ IMPORTANTE: NO usamos contains(subject,...) en $filter porque puede dar InefficientFilter.
    """
    base = f"{GRAPH}/{_user_segment()}/messages"
    params = {
        "$select": "id,subject,hasAttachments,categories,receivedDateTime,conversationId,isRead",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(max_messages, 500)),
    }

    if since_days is not None and since_days > 0:
        dt = datetime.now(timezone.utc) - timedelta(days=int(since_days))
        iso = dt.isoformat().replace("+00:00", "Z")
        params["$filter"] = f"receivedDateTime ge {iso}"

    # armamos query sin reventar commas/spaces innecesariamente
    q = "&".join([f"{k}={quote(v, safe='(),:$ ')}" for k, v in params.items()])
    url = f"{base}?{q}"

    data = _get_json_with_retries(url, retries=2, timeout=TIMEOUT)
    return (data or {}).get("value", [])


def buscar_mensajes_inbox_por_asunto(
    asunto_contiene: str,
    top: int = 50,
    since_days: int | None = 7,
    solo_con_adjuntos: bool = True,
):
    """
    Filtra por asunto en Python para evitar InefficientFilter en Graph.
    """
    asunto_contiene = (asunto_contiene or "").strip().lower()
    if not asunto_contiene:
        return []

    msgs = _listar_mensajes(max_messages=max(top, 20), since_days=since_days)

    out = []
    for m in msgs:
        subj = (m.get("subject") or "").lower()
        if asunto_contiene in subj:
            if solo_con_adjuntos and not m.get("hasAttachments"):
                continue
            out.append(m)
            if len(out) >= top:
                break
    return out


# ----------------------------
# Adjuntos
# ----------------------------
def _listar_adjuntos(msg_id: str):
    mid = quote(msg_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/messages/{mid}/attachments"
    data = _get_json_with_retries(url, retries=2, timeout=(15, 120))
    return (data or {}).get("value", [])


def _listar_adjuntos_zip(msg_id: str):
    items = _listar_adjuntos(msg_id)
    out = []
    for a in items:
        name = (a.get("name") or "")
        cty = (a.get("contentType") or "").lower()
        if name.lower().endswith(".zip") or "zip" in cty:
            out.append(a)
    return out


def listar_adjuntos_pdf(msg_id: str):
    items = _listar_adjuntos(msg_id)
    pdfs = []
    for a in items:
        name = (a.get("name") or "").lower()
        cty = (a.get("contentType") or "").lower()
        if name.endswith(".pdf") or "pdf" in cty:
            pdfs.append(a)
    return pdfs


def _descargar_adjunto(msg_id: str, att_id: str, dest_path: str):
    mid = quote(msg_id, safe="")
    aid = quote(att_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/messages/{mid}/attachments/{aid}"

    data = _get_json_with_retries(url, retries=2, timeout=(15, 120))
    if not data:
        print("[Graph] No se pudo obtener el adjunto (sin datos JSON).")
        return False

    content = data.get("contentBytes")
    if not content:
        print("[Graph] El adjunto no trae contentBytes.")
        return False

    try:
        raw = base64.b64decode(content)
    except Exception as e:
        print(f"[Graph] contentBytes inválido: {e}")
        return False

    try:
        with open(dest_path, "wb") as f:
            f.write(raw)
        return True
    except Exception as e:
        print(f"[FS] No se pudo escribir el adjunto en {dest_path}: {e}")
        return False


def descargar_adjunto_por_id(msg_id: str, att_id: str, dest_path: str) -> bool:
    return _descargar_adjunto(msg_id, att_id, dest_path)


# ------------------------------------
# ZIPs desde Inbox
# ------------------------------------
def listar_mensajes_zip_inbox(top: int = 300, since_days: int | None = None):
    msgs = _listar_mensajes(max_messages=min(top, 500), since_days=since_days)
    return [m for m in msgs if m.get("hasAttachments")]


def listar_adjuntos_zip(msg_id: str):
    return _listar_adjuntos_zip(msg_id)


def descargar_zips_validos(
    temp_check_dir,
    destino_dir,
    read_all=False,
    max_messages=200,
    since_days=None,
    required_categories=None,
):
    import zipfile

    os.makedirs(temp_check_dir, exist_ok=True)
    os.makedirs(destino_dir, exist_ok=True)

    msgs = _listar_mensajes(max_messages=max_messages, since_days=since_days)
    descargados = []

    for msg in msgs:
        if not msg.get("hasAttachments"):
            continue
        if not read_all and not _categorias_ok(msg, required_categories):
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

            tiene_xml = False
            try:
                with zipfile.ZipFile(tmp_path, "r") as zf:
                    tiene_xml = any(m.filename.lower().endswith(".xml") for m in zf.infolist())
            except Exception:
                tiene_xml = False

            if not tiene_xml:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
                continue

            dest_path = os.path.join(destino_dir, name)
            if not os.path.exists(dest_path):
                os.replace(tmp_path, dest_path)
                descargados.append(name)
            else:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    return descargados


# ------------------------------------------------
# PDFs desde carpeta de “Facturas aprobadas”
# ------------------------------------------------
def get_folder_id_by_name(root_display: str, name: str) -> str | None:
    inbox_url = f"{GRAPH}/{_user_segment()}/mailFolders/inbox"
    data = _get_json_with_retries(inbox_url, retries=2, timeout=TIMEOUT)
    if not data:
        return None
    root_id = data.get("id")
    if not root_id:
        return None

    childs_url = f"{GRAPH}/{_user_segment()}/mailFolders/{quote(root_id, safe='')}/childFolders?$top=500"
    data = _get_json_with_retries(childs_url, retries=2, timeout=TIMEOUT)
    if not data:
        return None

    for item in data.get("value", []):
        if (item.get("displayName") or "").strip().lower() == name.strip().lower():
            return item["id"]
    return None


def find_folder_id_anywhere(name: str) -> str | None:
    url = f"{GRAPH}/{_user_segment()}/mailFolders?$top=1000&$select=id,displayName"
    data = _get_json_with_retries(url, retries=2, timeout=TIMEOUT)
    if not data:
        return None
    for f in data.get("value", []):
        if (f.get("displayName") or "").strip().lower() == name.strip().lower():
            return f["id"]
    return None


def listar_mensajes_en_carpeta(folder_id: str, top: int = 200):
    fid = quote(folder_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/mailFolders/{fid}/messages"
    params = {
        "$select": "id,subject,hasAttachments,receivedDateTime,conversationId,isRead",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(top, 500)),
        "$filter": "isRead eq false",
    }
    q = "&".join([f"{k}={quote(v, safe='(),:$ ')}" for k, v in params.items()])
    data = _get_json_with_retries(f"{url}?{q}", retries=2, timeout=TIMEOUT)
    return (data or {}).get("value", [])


def marcar_mensaje_como_leido(msg_id: str) -> bool:
    try:
        mid = quote(msg_id, safe="")
        url = f"{GRAPH}/{_user_segment()}/messages/{mid}"
        payload = {"isRead": True}
        r = _SESSION.patch(
            url,
            headers=_h("application/json"),
            json=payload,
            timeout=TIMEOUT,
        )
        if r.status_code >= 400:
            try:
                body = r.json()
            except Exception:
                body = r.text
            print(f"[Graph] No se pudo marcar como leído ({r.status_code}): {body}")
            return False
        return True
    except Exception as e:
        print(f"[Graph] Error al marcar como leído: {e}")
        return False
