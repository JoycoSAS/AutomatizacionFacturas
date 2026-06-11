import sys
import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.m365.sp_graph import (
    upload_small_file,
    ensure_folder,
    SP_FOLDER as BASE_SP,
)

VERSION_UPLOAD = "2026-06-11-UPLOAD-CIERRE-DIARIO-SP-V1"

DATA_DIR = ROOT / "data"
CIERRES_DIR = DATA_DIR / "cierres_diarios"

NOW = datetime.datetime.now()
FECHA = NOW.strftime("%Y-%m-%d")
MES = NOW.strftime("%Y-%m")

CIERRE_DIA_DIR = CIERRES_DIR / FECHA
SP_CIERRE_DIR = f"{BASE_SP}/Backups/01_Cierres_Diarios/{MES}/{FECHA}".strip("/")

EXCLUIR_NOMBRES = {
    ".env",
}

EXCLUIR_EXT = {
    ".tmp",
    ".lock",
}


def debe_subir(path: Path) -> bool:
    if not path.is_file():
        return False

    if path.name.lower() in EXCLUIR_NOMBRES:
        return False

    if path.suffix.lower() in EXCLUIR_EXT:
        return False

    return True


def subir_archivo(local_path: Path, sp_dir: str) -> bool:
    sp_path = f"{sp_dir}/{local_path.name}".strip("/")

    print("☁️ Subiendo a SharePoint:")
    print(f"   Local: {local_path}")
    print(f"   SP:    {sp_path}")

    try:
        upload_small_file(str(local_path), sp_path, mode="replace")
        print(f"✅ Subido correctamente: {local_path.name}")
        return True
    except Exception as e:
        print(f"❌ Error subiendo {local_path.name}: {e}")
        return False


def main() -> int:
    print("=" * 100)
    print("SUBIDA CIERRE DIARIO A SHAREPOINT - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_UPLOAD}")
    print(f"Root: {ROOT}")
    print(f"Fecha: {FECHA}")
    print(f"Carpeta local: {CIERRE_DIA_DIR}")
    print(f"Carpeta SharePoint: {SP_CIERRE_DIR}")
    print("-" * 100)

    if not CIERRE_DIA_DIR.exists():
        print("❌ No existe carpeta de cierre diario local.")
        return 1

    archivos = [p for p in sorted(CIERRE_DIA_DIR.iterdir()) if debe_subir(p)]

    if not archivos:
        print("❌ No hay archivos válidos para subir.")
        return 1

    try:
        ensure_folder(SP_CIERRE_DIR)
        print("✅ Carpeta SharePoint verificada/creada.")
    except Exception as e:
        print(f"❌ No se pudo verificar/crear carpeta SharePoint: {e}")
        return 1

    ok_total = True

    for archivo in archivos:
        ok = subir_archivo(archivo, SP_CIERRE_DIR)
        if not ok:
            ok_total = False

    print("-" * 100)

    if ok_total:
        print("✅ Subida de cierre diario a SharePoint terminada correctamente.")
        print("=" * 100)
        return 0

    print("❌ Subida de cierre diario terminó con errores.")
    print("=" * 100)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
