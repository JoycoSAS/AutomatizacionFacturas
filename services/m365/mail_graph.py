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

# (connect, read) timeouts
TIMEOUT = (15, 60)
_SESSION = requests.Session()
_SESSION.headers.update({
    "Accept": "application/json"
})


# ----------------------------
# Helpers HTTP / Autenticación
# ----------------------------
def _h(content_type: str | None = None):
    """Headers con Bearer actual, opcional Content-Type."""
    h = {"Authorization": f"Bearer {get_access_token()}"}
    if content_type:
        h["Content-Type"] = content_type
    return h


def _user_segment() -> str:
    """Siempre usar /users/{mailbox}; evita /me. Codificamos solo el UPN."""
    return f"users/{quote(MAILBOX)}"


def _get(url, **kwargs):
    """GET simple con diagnóstico si HTTP >= 400 (no se usa para parseos frágiles)."""
    r = _SESSION.get(url, headers=_h(), timeout=kwargs.pop("timeout", TIMEOUT))
    if r.status_code >= 400:
        try:
            print("[Graph ERROR]", r.status_code, url, "->", r.json())
        except Exception:
            print("[Graph ERROR]", r.status_code, url, "->", r.text)
        r.raise_for_status()
    return r


def _get_json_with_retries(url: str, retries: int = 2, timeout=TIMEOUT):
    """
    GET + parse JSON con reintentos si la respuesta viene truncada/no-JSON.
    Devuelve dict (JSON) o None si no fue posible.
    """
    delay = 1.5
    for attempt in range(retries + 1):
        try:
            r = _SESSION.get(url, headers=_h(), timeout=timeout)
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
                    print(
                        f"[Graph] Respuesta no-JSON (posible problema de CONEXIÓN a Microsoft Graph/Azure): {e}"
                    )
                    snippet = (r.text[:200] + "...") if isinstance(r.text, str) and len(r.text) > 200 else r.text
                    if snippet:
                        print(f"→ Cuerpo parcial: {snippet}")
        except requests.RequestException as e:
            print(
                f"[Graph] Error de red/timeout al llamar a Graph "
                f"(posible inestabilidad de conexión o servicio en Azure): {e}"
            )

        if attempt < retries:
            print(f"[Graph] Reintentando ({attempt + 1}/{retries}) en {delay:.1f}s…")
            time.sleep(delay)
            delay *= 2

    return None


# ------------------------------------
# BLOQUE: ZIPs desde Inbox (flujo base)
# ------------------------------------
def _categorias_ok(msg, required_categories=None):
    if not required_categories:
        return True
    cats = set([c.lower() for c in (msg.get("categories") or [])])
    return all(c.lower() in cats for c in required_categories)


def _listar_mensajes(max_messages=200, since_days=None):
    """
    Lista mensajes del buzón (ordenados desc).
    Aplica filtro por fecha si since_days viene definido.
    """
    base = f"{GRAPH}/{_user_segment()}/messages"
    params = {
        "$select": "id,subject,hasAttachments,categories,receivedDateTime",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(max_messages, 500)),
    }

    # Filtro por fecha (UTC)
    if since_days is not None and since_days > 0:
        dt = datetime.now(timezone.utc) - timedelta(days=int(since_days))
        iso = dt.isoformat().replace("+00:00", "Z")
        params["$filter"] = f"receivedDateTime ge {iso}"

    url = base + "?" + "&".join([f"{k}={quote(v)}" for k, v in params.items()])
    data = _get_json_with_retries(url, retries=2, timeout=TIMEOUT)
    return (data or {}).get("value", [])


def _listar_adjuntos(msg_id: str):
    """
    Lista adjuntos de un mensaje dado.
    Reintenta si la respuesta de Graph viene corrupta o no es JSON válido.
    """
    mid = quote(msg_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/messages/{mid}/attachments"
    data = _get_json_with_retries(url, retries=2, timeout=(15, 120))
    return (data or {}).get("value", [])


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
    """
    Descarga el adjunto usando GET /attachments/{id} (lee contentBytes).
    Si el adjunto no trae contentBytes, devuelve False.
    """
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


def descargar_zips_validos(
    temp_check_dir,
    destino_dir,
    read_all=False,
    max_messages=200,
    since_days=None,
    required_categories=None,
):
    """
    Barre mensajes recientes, descarga ZIPs con contentBytes y que contengan XML,
    y los mueve a destino si no existen ya.
    """
    import zipfile
    import os

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

            # ¿El ZIP trae XML?
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
# BLOQUE: PDFs desde carpeta de “Facturas aprobadas”
# ------------------------------------------------
def get_folder_id_by_name(root_display: str, name: str) -> str | None:
    """Busca 'name' como hija directa de Inbox."""
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
    """Búsqueda global por displayName en todo el buzón."""
    url = f"{GRAPH}/{_user_segment()}/mailFolders?$top=1000&$select=id,displayName"
    data = _get_json_with_retries(url, retries=2, timeout=TIMEOUT)
    if not data:
        return None
    for f in data.get("value", []):
        if (f.get("displayName") or "").strip().lower() == name.strip().lower():
            return f["id"]
    return None


def listar_mensajes_en_carpeta(folder_id: str, top: int = 200):
    """
    Lista mensajes (desc) de una carpeta específica (por id).

    🔧 Solo devuelve mensajes NO leídos (isRead = false) para evitar
    reprocesar aprobaciones ya tratadas.
    """
    fid = quote(folder_id, safe="")
    url = f"{GRAPH}/{_user_segment()}/mailFolders/{fid}/messages"
    params = {
        "$select": "id,subject,hasAttachments,receivedDateTime,conversationId,isRead",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(top, 500)),
        "$filter": "isRead eq false",
    }
    q = "&".join([f"{k}={quote(v)}" for k, v in params.items()])
    data = _get_json_with_retries(f"{url}?{q}", retries=2, timeout=TIMEOUT)
    return (data or {}).get("value", [])


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
    try:
        data = base64.b64decode(content)
    except Exception:
        return False
    with open(dest_path, "wb") as f:
        f.write(data)
    return True


def listar_mensajes_zip_inbox(top: int = 300, since_days: int | None = None):
    """
    Lista mensajes del Inbox (desc) con adjuntos; opcionalmente filtra por fecha
    usando $filter sobre receivedDateTime >= now-<since_days>.
    """
    base = f"{GRAPH}/{_user_segment()}/messages"
    params = {
        "$select": "id,subject,hasAttachments,receivedDateTime",
        "$orderby": "receivedDateTime desc",
        "$top": str(min(top, 500)),
    }

    if since_days is not None and since_days > 0:
        dt = datetime.now(timezone.utc) - timedelta(days=int(since_days))
        iso = dt.isoformat().replace("+00:00", "Z")
        params["$filter"] = f"receivedDateTime ge {iso}"

    url = base + "?" + "&".join([f"{k}={quote(v)}" for k, v in params.items()])
    data = _get_json_with_retries(url, retries=2, timeout=TIMEOUT)
    values = (data or {}).get("value", [])
    return [m for m in values if m.get("hasAttachments")]


def listar_adjuntos_zip(msg_id: str):
    """Lista SOLO los ZIPs; descarga con _descargar_adjunto()."""
    return _listar_adjuntos_zip(msg_id)


# API pública para descargar por id (útil desde controladores)
def descargar_adjunto_por_id(msg_id: str, att_id: str, dest_path: str) -> bool:
    return _descargar_adjunto(msg_id, att_id, dest_path)


# 🔧 NUEVO: marcar un mensaje como leído
def marcar_mensaje_como_leido(msg_id: str) -> bool:
    """
    Marca un mensaje como leído (isRead=true).
    Se usa solo cuando YA se procesó con éxito la factura asociada.
    """
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
