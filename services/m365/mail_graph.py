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


def _get_json_with_retries(
    url: str,
    retries: int = 2,
    timeout=TIMEOUT,
    headers_extra: dict | None = None
):
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


def _env_int(name: str, default: int | None = None) -> int | None:
    raw = (os.getenv(name) or "").strip()
    if raw.isdigit():
        return int(raw)
    return default


def _env_bool(name: str, default: bool = False) -> bool:
    raw = (os.getenv(name) or "").strip().lower()
    if raw in ("1", "true", "yes", "y", "si", "sí", "on"):
        return True
    if raw in ("0", "false", "no", "n", "off"):
        return False
    return default


# ----------------------------
# Listar mensajes (Inbox)
# ----------------------------
def _listar_mensajes(max_messages=200, since_days=None):
    """
    Lista mensajes del Inbox (ordenados desc).
    ✅ IMPORTANTE: NO usamos contains(subject,...) en $filter porque puede dar InefficientFilter.
    """
    max_env = _env_int("MAX_MESSAGES", 200) or 200
    max_messages = min(int(max_messages or 200), 500, max_env)

    base = f"{GRAPH}/{_user_segment()}/messages"
    params = {
        "$select": "id,subject,hasAttachments,categories,receivedDateTime,conversationId,isRead,bodyPreview",
        "$orderby": "receivedDateTime desc",
        "$top": str(max_messages),
    }

    if since_days is not None and since_days > 0:
        dt = datetime.now(timezone.utc) - timedelta(days=int(since_days))
        iso = dt.isoformat().replace("+00:00", "Z")
        params["$filter"] = f"receivedDateTime ge {iso}"

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

    ✅ CORREGIDO:
    - Si asunto_contiene viene vacío, devolvemos mensajes sin filtrar (para tus fallbacks).
    """
    top = int(top or 50)
    if top < 1:
        return []

    if not (asunto_contiene or "").strip():
        msgs = _listar_mensajes(max_messages=max(top, 20), since_days=since_days)
        if solo_con_adjuntos:
            msgs = [m for m in msgs if m.get("hasAttachments")]
        return msgs[:top]

    asunto_contiene = asunto_contiene.strip().lower()

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


def listar_mensajes_en_carpeta(
    folder_id: str,
    top: int = 200,
    unread_only: bool | None = None,
    since_days: int | None = None,
):
    fid = quote(folder_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/mailFolders/{fid}/messages"

    if unread_only is None:
        unread_only = _env_bool("MAIL_UNREAD_ONLY", True)

    if since_days is None:
        since_days = _env_int("MAIL_LOOKBACK_DAYS", None)

    max_env = _env_int("MAX_MESSAGES", 200) or 200
    top = min(int(top or 200), 500, max_env)

    params = {
        "$select": "id,subject,hasAttachments,receivedDateTime,conversationId,isRead,bodyPreview",
        "$orderby": "receivedDateTime desc",
        "$top": str(top),
    }

    filtros = []
    if since_days is not None and since_days > 0:
        dt = datetime.now(timezone.utc) - timedelta(days=int(since_days))
        iso = dt.isoformat().replace("+00:00", "Z")
        filtros.append(f"receivedDateTime ge {iso}")

    if unread_only:
        filtros.append("isRead eq false")

    if filtros:
        params["$filter"] = " and ".join(filtros)

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