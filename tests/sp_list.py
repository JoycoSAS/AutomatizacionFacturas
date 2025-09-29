import os
from urllib.parse import quote
from services.m365.sp_graph import _SESSION, _h, GRAPH, DRIVE_ID, SP_FOLDER

def ls(path):
    url = f"{GRAPH}/drives/{DRIVE_ID}/root:/{quote(path)}:/children"
    r = _SESSION.get(url, headers=_h(), timeout=(15,60))
    if r.status_code == 404:
        print("No existe:", path)
        return
    r.raise_for_status()
    items = r.json().get("value", [])
    print(f"[{path}]")
    for it in items:
        mark = "/" if "folder" in it else ""
        print(" -", it["name"] + mark)

if __name__ == "__main__":
    base = SP_FOLDER
    ls(base)
    ls(f"{base}/excel")
    from datetime import datetime
    today = datetime.now().strftime("%Y-%m-%d")
    ls(f"{base}/adjuntos/{today}")
    ls(f"{base}/extraidos/{today}")
