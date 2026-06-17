# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Backup mensual local seguro

Objetivo:
- Generar un backup mensual local en ZIP.
- Incluir Excel operativo, auditorías, logs relevantes, state/AIDX y configuración redactada.
- Generar manifest con hash SHA256 por archivo.
- NO subir a SharePoint. La subida se hace con scripts/subir_backup_mensual_sharepoint.py.
- NO incluir .env real ni secretos.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import config  # noqa: F401
except Exception:
    # El backup local no debe depender de que config imprima o cargue perfecto.
    pass

VERSION_BACKUP_MENSUAL = "2026-06-17-BACKUP-MENSUAL-SEGURO-V2-LOCAL-RESTAURADO"

DATA_DIR = ROOT / "data"
AUDIT_DIR = DATA_DIR / "audit"
STATE_DIR = DATA_DIR / "state"
LOGS_DIR = ROOT / "logs"
DATA_LOGS_DIR = DATA_DIR / "logs"
BACKUPS_MENSUALES_DIR = DATA_DIR / "backups_mensuales"

NOW = _dt.datetime.now()
MES = NOW.strftime("%Y-%m")
MES_FILE = NOW.strftime("%Y_%m")
STAMP = NOW.strftime("%Y%m%d_%H%M%S")

BACKUP_MES_DIR = BACKUPS_MENSUALES_DIR / MES
ZIP_NAME = f"backup_mensual_{MES_FILE}_{STAMP}.zip"
MANIFEST_NAME = f"manifest_backup_mensual_{MES_FILE}_{STAMP}.json"
RESUMEN_NAME = f"RESUMEN_BACKUP_MENSUAL_{MES}_{STAMP}.txt"

ZIP_PATH = BACKUP_MES_DIR / ZIP_NAME
MANIFEST_PATH = BACKUP_MES_DIR / MANIFEST_NAME
RESUMEN_PATH = BACKUP_MES_DIR / RESUMEN_NAME

SECRET_PATTERNS = (
    "SECRET",
    "PASSWORD",
    "PASS",
    "TOKEN",
    "CLIENT_SECRET",
    "TENANT_ID",
    "CLIENT_ID",
    "AUTHORITY",
)

EXCLUIR_DIRS = {
    "__pycache__",
    ".git",
    ".venv",
    "venv",
    "env",
    "temp",
    "tmp",
    "adjuntos",
    "extraidos",
    "temp_check",
    "cierres_diarios",
    "backups_mensuales",
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def normalizar_rel(path: Path) -> str:
    return path.as_posix().replace("\\", "/")


def redactar_env_line(line: str) -> str:
    raw = line.rstrip("\n\r")
    if not raw or raw.lstrip().startswith("#") or "=" not in raw:
        return raw
    key, value = raw.split("=", 1)
    key_upper = key.strip().upper()
    if any(p in key_upper for p in SECRET_PATTERNS):
        return f"{key}=***REDACTADO***"
    # También redacción defensiva para valores con apariencia de secreto largo.
    if re.search(r"[A-Za-z0-9_\-]{32,}", value) and key_upper not in {
        "SP_DRIVE_ID",
        "SP_DRIVE_ID_RADICADOS",
        "SP_BACKUP2_DRIVE_ID",
    }:
        return f"{key}=***REDACTADO***"
    return raw


def crear_snapshot_env_redactado(work_dir: Path) -> Optional[Path]:
    env_path = ROOT / ".env"
    if not env_path.exists():
        return None
    destino_dir = work_dir / "05_config_redactada"
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / f"snapshot_env_redactado_{MES}_{STAMP}.txt"
    try:
        txt = env_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception:
        txt = env_path.read_text(errors="replace").splitlines()
    destino.write_text("\n".join(redactar_env_line(x) for x in txt) + "\n", encoding="utf-8")
    return destino


def copiar_si_existe(origen: Path, destino_rel: str, work_dir: Path, archivos: List[Tuple[Path, str]]) -> None:
    if not origen.exists() or not origen.is_file():
        return
    destino = work_dir / destino_rel
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origen, destino)
    archivos.append((destino, normalizar_rel(Path(destino_rel))))


def incluir_patron(base: Path, patron: str, destino_base: str, work_dir: Path, archivos: List[Tuple[Path, str]]) -> None:
    if not base.exists():
        return
    for p in sorted(base.glob(patron)):
        if not p.is_file():
            continue
        rel = Path(destino_base) / p.name
        copiar_si_existe(p, normalizar_rel(rel), work_dir, archivos)


def incluir_logs(base: Path, destino_base: str, work_dir: Path, archivos: List[Tuple[Path, str]]) -> None:
    if not base.exists():
        return
    for p in sorted(base.rglob("*")):
        if not p.is_file():
            continue
        partes = {x.lower() for x in p.relative_to(base).parts}
        if partes & EXCLUIR_DIRS:
            continue
        # Evita subir logs enormes o binarios raros; los logs esperados son texto/csv/json.
        if p.suffix.lower() not in {".log", ".txt", ".csv", ".json"}:
            continue
        rel = Path(destino_base) / p.relative_to(base)
        copiar_si_existe(p, normalizar_rel(rel), work_dir, archivos)


def crear_manifest(archivos: List[Tuple[Path, str]]) -> Dict[str, object]:
    items = []
    total_bytes = 0
    for path, rel in archivos:
        size = path.stat().st_size
        total_bytes += size
        items.append(
            {
                "ruta_relativa": rel,
                "nombre": path.name,
                "bytes": size,
                "sha256": sha256_file(path),
                "modificado_local": _dt.datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds"),
            }
        )
    return {
        "tipo": "backup_mensual_local",
        "proyecto": "Automatizacion de Facturas JOYCO",
        "version_script": VERSION_BACKUP_MENSUAL,
        "generado_en": NOW.isoformat(timespec="seconds"),
        "mes": MES,
        "root": str(ROOT),
        "zip": ZIP_NAME,
        "total_archivos": len(items),
        "total_bytes": total_bytes,
        "items": items,
        "nota_seguridad": "El archivo .env real no se incluye. Solo se incluye snapshot redactado si existe.",
    }


def escribir_resumen(manifest: Dict[str, object]) -> None:
    lines = [
        "BACKUP MENSUAL LOCAL SEGURO - FACTURAS JOYCO",
        "=" * 80,
        f"Versión: {VERSION_BACKUP_MENSUAL}",
        f"Fecha generación: {manifest['generado_en']}",
        f"Mes: {MES}",
        f"Root: {ROOT}",
        f"Carpeta destino: {BACKUP_MES_DIR}",
        f"ZIP: {ZIP_PATH}",
        f"Manifest: {MANIFEST_PATH}",
        f"Total archivos incluidos: {manifest['total_archivos']}",
        f"Total bytes incluidos: {manifest['total_bytes']}",
        "",
        "Contenido principal incluido:",
        "- Excel operativo local y/o historial si existen.",
        "- Auditorías CSV disponibles.",
        "- State/AIDX/ProcessedStore.",
        "- Logs relevantes si existen.",
        "- Snapshot .env redactado, sin secretos.",
        "",
        "Regla de seguridad:",
        "- No se incluye .env real ni secretos.",
        "- No se incluyen carpetas temporales, adjuntos pesados, extraídos, cierres diarios ni backups mensuales previos.",
        "=" * 80,
    ]
    RESUMEN_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def generar_backup() -> int:
    print("=" * 100)
    print("BACKUP MENSUAL LOCAL SEGURO - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_BACKUP_MENSUAL}")
    print(f"Root: {ROOT}")
    print(f"Mes: {MES}")
    print(f"Carpeta destino: {BACKUP_MES_DIR}")
    print("-" * 100)

    BACKUP_MES_DIR.mkdir(parents=True, exist_ok=True)
    work_dir = BACKUP_MES_DIR / f"_work_backup_mensual_{STAMP}"
    if work_dir.exists():
        shutil.rmtree(work_dir, ignore_errors=True)
    work_dir.mkdir(parents=True, exist_ok=True)

    archivos: List[Tuple[Path, str]] = []

    try:
        # Excel operativo e historial.
        copiar_si_existe(DATA_DIR / "facturas.xlsx", "01_excel/facturas.xlsx", work_dir, archivos)
        copiar_si_existe(DATA_DIR / "historial_ejecuciones.xlsx", "01_excel/historial_ejecuciones.xlsx", work_dir, archivos)

        # Auditorías: soporta tanto data/audit como archivos audit_* en data root.
        incluir_patron(AUDIT_DIR, "audit_*.csv", "02_auditoria", work_dir, archivos)
        incluir_patron(DATA_DIR, "audit_*.csv", "02_auditoria", work_dir, archivos)

        # State crítico.
        copiar_si_existe(STATE_DIR / "processed_messages.json", "03_state/processed_messages.json", work_dir, archivos)
        copiar_si_existe(STATE_DIR / "attachment_index_store.json", "03_state/attachment_index_store.json", work_dir, archivos)
        copiar_si_existe(STATE_DIR / "attachment_index_seen_messages.json", "03_state/attachment_index_seen_messages.json", work_dir, archivos)

        # Logs si existen.
        incluir_logs(LOGS_DIR, "04_logs", work_dir, archivos)
        incluir_logs(DATA_LOGS_DIR, "04_logs_data", work_dir, archivos)

        # Config redactada.
        env_redactado = crear_snapshot_env_redactado(work_dir)
        if env_redactado:
            archivos.append((env_redactado, normalizar_rel(Path("05_config_redactada") / env_redactado.name)))

        if not archivos:
            print("❌ No se encontraron archivos para incluir en el backup mensual.")
            return 1

        manifest = crear_manifest(archivos)
        MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        escribir_resumen(manifest)

        # También incluir manifest y resumen dentro del ZIP para que el ZIP sea autocontenido.
        archivos_zip = list(archivos)
        archivos_zip.append((MANIFEST_PATH, MANIFEST_NAME))
        archivos_zip.append((RESUMEN_PATH, RESUMEN_NAME))

        if ZIP_PATH.exists():
            ZIP_PATH.unlink()
        with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
            for path, rel in archivos_zip:
                zf.write(path, arcname=rel)

        zip_hash = sha256_file(ZIP_PATH)
        manifest["zip_bytes"] = ZIP_PATH.stat().st_size
        manifest["zip_sha256"] = zip_hash
        MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

        print(f"✅ Archivos incluidos: {manifest['total_archivos']}")
        print(f"✅ ZIP generado: {ZIP_PATH}")
        print(f"✅ ZIP bytes: {ZIP_PATH.stat().st_size}")
        print(f"✅ ZIP SHA256: {zip_hash}")
        print(f"✅ Manifest: {MANIFEST_PATH}")
        print(f"✅ Resumen: {RESUMEN_PATH}")
        print("=" * 100)
        print("✅ Backup mensual local generado correctamente.")
        print("Siguiente paso: python scripts\\subir_backup_mensual_sharepoint.py")
        print("=" * 100)
        return 0

    except Exception as exc:
        print(f"❌ Error generando backup mensual local: {exc}")
        print("⚠️ No subas ni archives nada hasta revisar el error.")
        return 1
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(generar_backup())
