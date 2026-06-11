import os
import json
import zipfile
import hashlib
import datetime
from pathlib import Path

VERSION_BACKUP = "2026-06-11-BACKUP-MENSUAL-SEGURO-V1"

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"

CIERRES_DIR = DATA_DIR / "cierres_diarios"
BACKUPS_DIR = DATA_DIR / "backups_mensuales"

NOW = datetime.datetime.now()
MES = NOW.strftime("%Y-%m")
MES_FILE = NOW.strftime("%Y_%m")
STAMP = NOW.strftime("%Y%m%d_%H%M%S")

DESTINO_DIR = BACKUPS_DIR / MES
ZIP_PATH = DESTINO_DIR / f"backup_mensual_{MES_FILE}_{STAMP}.zip"
MANIFEST_PATH = DESTINO_DIR / f"manifest_backup_mensual_{MES_FILE}_{STAMP}.json"

EXCLUIR_NOMBRES = {
    ".env",
}

EXCLUIR_EXTENSIONES = {
    ".tmp",
    ".lock",
}

def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def debe_excluir(path: Path) -> bool:
    name = path.name.lower().strip()

    if name in EXCLUIR_NOMBRES:
        return True

    if path.suffix.lower() in EXCLUIR_EXTENSIONES:
        return True

    return False

def obtener_cierres_del_mes() -> list[Path]:
    if not CIERRES_DIR.exists():
        return []

    carpetas = []
    for p in CIERRES_DIR.iterdir():
        if not p.is_dir():
            continue
        if p.name.startswith(MES):
            carpetas.append(p)

    return sorted(carpetas)

def recolectar_archivos(carpetas: list[Path]) -> list[Path]:
    archivos = []

    for carpeta in carpetas:
        for p in carpeta.rglob("*"):
            if not p.is_file():
                continue
            if debe_excluir(p):
                continue
            archivos.append(p)

    return sorted(archivos)

def main() -> int:
    print("=" * 100)
    print("BACKUP MENSUAL SEGURO - AUTOMATIZACIÓN FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_BACKUP}")
    print(f"Root: {ROOT}")
    print(f"Mes: {MES}")
    print(f"Destino ZIP: {ZIP_PATH}")
    print("-" * 100)

    carpetas_mes = obtener_cierres_del_mes()

    if not carpetas_mes:
        print(f"⚠️ No hay cierres diarios para el mes {MES}. No se genera backup.")
        return 2

    print(f"Carpetas de cierre encontradas: {len(carpetas_mes)}")
    for c in carpetas_mes:
        print(f"  - {c}")

    archivos = recolectar_archivos(carpetas_mes)

    if not archivos:
        print("⚠️ Hay carpetas de cierre, pero no hay archivos respaldables.")
        return 2

    DESTINO_DIR.mkdir(parents=True, exist_ok=True)

    manifest = {
        "version": VERSION_BACKUP,
        "fecha_generacion": NOW.isoformat(timespec="seconds"),
        "mes": MES,
        "root": str(ROOT),
        "zip_path": str(ZIP_PATH),
        "carpetas_cierre": [str(c) for c in carpetas_mes],
        "archivos_total": len(archivos),
        "archivos": [],
    }

    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for src in archivos:
            rel = src.relative_to(DATA_DIR)
            arcname = str(Path("data") / rel).replace("\\", "/")

            zf.write(src, arcname)

            item = {
                "origen": str(src),
                "zip_name": arcname,
                "size_bytes": src.stat().st_size,
                "sha256": sha256_file(src),
            }
            manifest["archivos"].append(item)

            print(f"✅ Agregado: {arcname}")

    manifest["zip_size_bytes"] = ZIP_PATH.stat().st_size
    manifest["zip_sha256"] = sha256_file(ZIP_PATH)

    MANIFEST_PATH.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print("-" * 100)
    print(f"✅ ZIP generado: {ZIP_PATH}")
    print(f"✅ Manifest generado: {MANIFEST_PATH}")
    print(f"Archivos incluidos: {len(archivos)}")
    print(f"Tamaño ZIP: {manifest['zip_size_bytes']} bytes")
    print("=" * 100)

    return 0

if __name__ == "__main__":
    raise SystemExit(main())
