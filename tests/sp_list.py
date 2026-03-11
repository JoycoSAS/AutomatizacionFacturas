import os
from urllib.parse import quote

from services.m365.sp_graph import _SESSION, _h, GRAPH

DRIVE_ID = (os.getenv("SP_DRIVE_ID") or "").strip()
# ✅ Permitir SP_FOLDER vacío para listar root
SP_FOLDER = (os.getenv("SP_FOLDER") or "").strip()


if not DRIVE_ID:
    raise SystemExit("❌ Falta SP_DRIVE_ID en el .env")


def _normalize(path: str) -> str:
    """Normaliza rutas tipo SharePoint: quita espacios extremos y slashes repetidos."""
    path = (path or "").strip()
    path = path.strip("/")
    return path


def _children_url(path: str) -> str:
    """
    Construye el URL de Graph para listar children.
    - Si path está vacío => root/children
    - Si no => root:/{path}:/children
    """
    path = _normalize(path)
    if not path:
        return f"{GRAPH}/drives/{DRIVE_ID}/root/children"
    return f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(path)}:/children"


def ls(path: str):
    path_norm = _normalize(path)
    url = _children_url(path_norm)
    r = _SESSION.get(url, headers=_h(), timeout=(15, 60))

    if r.status_code == 404:
        print(f"❌ No existe: {path_norm or '/'}")
        return []

    r.raise_for_status()
    items = r.json().get("value", [])

    print(f"\n📂 [{path_norm or '/'}] ({len(items)} items)")
    for it in items:
        if "folder" in it:
            print(f"  📁 {it['name']}/")
        else:
            print(f"  📄 {it['name']}")

    return items


def list_subfolders(base: str, max_depth: int = 1):
    """
    Lista recursivamente subcarpetas hasta max_depth.
    max_depth=1 => lista solo hijos directos (tu comportamiento actual).
    """
    base_norm = _normalize(base)
    items = ls(base_norm)

    if max_depth <= 0:
        return

    for it in items:
        if "folder" in it:
            sub = f"{base_norm}/{it['name']}" if base_norm else it["name"]
            if max_depth == 1:
                # Solo un nivel (como tu script original)
                ls(sub)
            else:
                list_subfolders(sub, max_depth=max_depth - 1)


if __name__ == "__main__":
    # ✅ 1) Listar carpeta base (si SP_FOLDER está vacío => root)
    # ✅ 2) Listar subcarpetas automáticamente (1 nivel, como tenías)
    list_subfolders(SP_FOLDER, max_depth=1)