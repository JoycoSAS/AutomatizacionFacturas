import os
from pathlib import Path
from dotenv import load_dotenv

from config import RADICADOS_SP_RELATIVE_PATH, RADICADOS_LOCAL_PATH
from services.m365.sp_graph import download_small_file

load_dotenv()

def main():
    drive_id = (os.getenv("SP_DRIVE_ID_RADICADOS") or "").strip()

    print("DRIVE_RADICADOS:", drive_id)
    print("SP_PATH:", RADICADOS_SP_RELATIVE_PATH)
    print("LOCAL :", RADICADOS_LOCAL_PATH)

    # asegurar carpeta local
    Path(RADICADOS_LOCAL_PATH).parent.mkdir(parents=True, exist_ok=True)

    ok = download_small_file(
        RADICADOS_SP_RELATIVE_PATH,
        RADICADOS_LOCAL_PATH,
        drive_id=drive_id
    )

    print("✅ Descarga OK?", ok)

if __name__ == "__main__":
    main()
