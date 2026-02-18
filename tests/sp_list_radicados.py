import os
from dotenv import load_dotenv
from services.m365.sp_graph import list_children

load_dotenv()

def main():
    drive2 = os.getenv("SP_DRIVE_ID_RADICADOS")
    folder2 = (os.getenv("SP_FOLDER_RADICADOS") or "").strip().strip("/")

    if not drive2:
        print("❌ Falta SP_DRIVE_ID_RADICADOS en .env")
        return

    print("=== SHAREPOINT 2 (RADICADOS) ===")
    print("DRIVE:", drive2)
    print("PATH :", folder2 or "/")

    try:
        items = list_children(rel_path=folder2, drive_id=drive2, top=999)
    except Exception as e:
        print("❌ Error listando children:", e)
        return

    print(f"\n📂 [{folder2 or '/'}] ({len(items)} items)")
    for it in items:
        name = it.get("name")
        is_folder = "folder" in it
        print(("  📁 " if is_folder else "  📄 ") + str(name))

if __name__ == "__main__":
    main()
