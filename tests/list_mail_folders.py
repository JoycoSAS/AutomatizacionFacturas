# tests/list_mail_folders.py
from services.m365.mail_graph import _SESSION, _h, _user_segment, GRAPH

if __name__ == "__main__":
    url = f"{GRAPH}/{_user_segment()}/mailFolders?$top=1000&$select=id,displayName,parentFolderId"
    r = _SESSION.get(url, headers=_h(), timeout=(15,60)); r.raise_for_status()
    for f in r.json().get("value", []):
        print(f"{f['displayName']} | id={f['id']} | parent={f.get('parentFolderId')}")
