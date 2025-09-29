# services/m365/mail_graph.py
import os
import io
import time
import zipfile
import datetime as dt
from typing import Iterator, List, Tuple

import requests
from dotenv import load_dotenv

from .token import get_access_token

load_dotenv()

GRAPH = "https://graph.microsoft.com/v1.0"
MAIL_USER = (
    os.getenv("MAILBOX_USER")
    or os.getenv("MAIL_USER")
    or "auxiliar.infraestructura@joyco.com.co"
)
TIMEOUT = (15, 60)

_SESSION = requests.Session()


def _h():
    return {"Authorization": f"Bearer {get_access_token()}"}


def _req(call, max_retries=4):
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
        print(f"[Graph GET ERROR] {r.status_code} {r.request.method} {r.url} {body}")
        r.raise_for_status()


def _iter_messages_with_attachments(
    read_all: bool = False,
    max_messages: int = 200,
    since_days: int | None = None,
) -> Iterator[dict]:
    """
    Itera mensajes del Inbox ordenados desc por fecha.
    No usamos $filter para evitar 'InefficientFilter'; filtramos en cliente.
    """
    params = {
        "$select": "id,subject,receivedDateTime,hasAttachments",
        "$orderby": "receivedDateTime desc",
        "$top": 50,  # paginado
    }

    base_url = f"{GRAPH}/users/{MAIL_USER}/mailFolders/Inbox/messages"
    fetched = 0
    cutoff = None
    if since_days and since_days > 0:
        cutoff = dt.datetime.utcnow() - dt.timedelta(days=since_days)

    next_url = base_url
    while next_url and fetched < max_messages:
        r = _req(lambda: _SESSION.get(next_url, headers=_h(), params=params, timeout=TIMEOUT))
        data = r.json()

        for msg in data.get("value", []):
            if not msg.get("hasAttachments"):
                continue

            if cutoff:
                rdt = msg.get("receivedDateTime")
                try:
                    rec = dt.datetime.fromisoformat(rdt.replace("Z", "+00:00"))
                    if rec.tzinfo:
                        rec = rec.astimezone(dt.timezone.utc).replace(tzinfo=None)
                except Exception:
                    rec = None
                if rec and rec < cutoff:
                    return

            yield msg
            fetched += 1
            if fetched >= max_messages:
                return

        next_url = data.get("@odata.nextLink")
        params = None  # nextLink ya incluye los params


def _download_attachment_value(message_id: str, attachment_id: str) -> bytes:
    """
    Descarga el adjunto como binario usando /$value (forma soportada y estable).
    """
    url = f"{GRAPH}/users/{MAIL_USER}/messages/{message_id}/attachments/{attachment_id}/$value"
    r = _req(lambda: _SESSION.get(url, headers=_h(), timeout=(15, 300), stream=True))
    bio = io.BytesIO()
    for chunk in r.iter_content(1024 * 1024):
        if chunk:
            bio.write(chunk)
    return bio.getvalue()


def _iter_zip_attachments(message_id: str) -> Iterator[Tuple[str, bytes]]:
    """
    Devuelve (name, content_bytes) de adjuntos ZIP en un mensaje.
    1) Listamos adjuntos (sin @odata.type ni contentBytes en $select).
    2) Para cada .zip, descargamos el binario con /$value.
    """
    url = f"{GRAPH}/users/{MAIL_USER}/messages/{message_id}/attachments"
    params = {"$select": "id,name,contentType,size", "$top": 50}

    while url:
        r = _req(lambda: _SESSION.get(url, headers=_h(), params=params, timeout=TIMEOUT))
        data = r.json()
        for att in data.get("value", []):
            name = (att.get("name") or "").strip()
            if not name.lower().endswith(".zip"):
                continue
            att_id = att.get("id")
            if not att_id:
                continue

            try:
                content = _download_attachment_value(message_id, att_id)
            except Exception:
                continue

            yield name, content

        url = data.get("@odata.nextLink")
        params = None


def _zip_has_xml(content: bytes) -> bool:
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as z:
            return any(n.lower().endswith(".xml") for n in z.namelist())
    except Exception:
        return False


def _save(path: str, content: bytes):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(content)


def descargar_zips_validos(
    temp_check_dir: str,
    destino_dir: str,
    read_all: bool = False,
    max_messages: int = 200,
    since_days: int | None = None,
) -> List[str]:
    """
    Descarga ZIPs válidos (con al menos un XML) a temp_check_dir,
    y los mueve a destino_dir. Devuelve la lista de nombres guardados.
    """
    os.makedirs(temp_check_dir, exist_ok=True)
    os.makedirs(destino_dir, exist_ok=True)

    guardados: List[str] = []

    for msg in _iter_messages_with_attachments(
        read_all=read_all, max_messages=max_messages, since_days=since_days
    ):
        mid = msg["id"]
        for name, content in _iter_zip_attachments(mid):
            if name in guardados:
                continue
            if not _zip_has_xml(content):
                continue

            tmp_path = os.path.join(temp_check_dir, name)
            _save(tmp_path, content)

            final_path = os.path.join(destino_dir, name)
            if not os.path.exists(final_path):
                os.replace(tmp_path, final_path)
            else:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

            guardados.append(name)

    return guardados
