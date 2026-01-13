import os
from urllib.parse import quote

from services.m365.sp_graph import _SESSION, _h, GRAPH

DRIVE_ID = os.getenv("SP_DRIVE_ID")
SP_FOLDER = os.getenv("SP_FOLDER")

if not DRIVE_ID:
    raise SystemExit("❌ Falta SP_DRIVE_ID en el .env")
if not SP_FOLDER:
    raise SystemExit("❌ Falta SP_FOLDER en el .env")


def ls(path: str):
    url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(path)}:/children"
    r = _SESSION.get(url, headers=_h(), timeout=(15, 60))

    if r.status_code == 404:
        print(f"❌ No existe: {path}")
        return

    r.raise_for_status()
    items = r.json().get("value", [])

    print(f"\n📂 [{path}] ({len(items)} items)")
    for it in items:
        if "folder" in it:
            print(f"  📁 {it['name']}/")
        else:
            print(f"  📄 {it['name']}")


if __name__ == "__main__":
    # 1️⃣ Listar carpeta base
    ls(SP_FOLDER)

    # 2️⃣ Listar TODO lo que haya dentro automáticamente
    url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(SP_FOLDER)}:/children"
    r = _SESSION.get(url, headers=_h(), timeout=(15, 60))
    r.raise_for_status()

    for item in r.json().get("value", []):
        if "folder" in item:
            sub = f"{SP_FOLDER}/{item['name']}"
            ls(sub)
