import datetime
import hashlib
import json
import os
import platform
import shutil
import socket
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import config  # noqa: F401  # Carga .env/configuración del proyecto sin imprimir secretos.
except Exception:
    pass

VERSION_CIERRE = "2026-06-12-CIERRE-DIARIO-SEGURO-V4-LOCAL-MKDIR-CONFIG"

DATA_DIR = ROOT / "data"
STATE_DIR = DATA_DIR / "state"
LOGS_DIR = DATA_DIR / "logs"
CIERRES_DIR = DATA_DIR / "cierres_diarios"
LOCKS_DIR = DATA_DIR / "locks"

NOW = datetime.datetime.now()
FECHA = NOW.strftime("%Y-%m-%d")
MES = NOW.strftime("%Y-%m")
STAMP = NOW.strftime("%Y%m%d_%H%M%S")

CIERRE_DIA_DIR = CIERRES_DIR / FECHA
LOCK_FILE = LOCKS_DIR / "cierre_diario_seguro.lock"
LOCK_TTL_SECONDS = int(os.getenv("LOCK_TTL_SECONDS", "3600") or "3600")

NOMBRES_ESTADO = [
    "processed_messages.json",
    "attachment_index_store.json",
    "attachment_index_seen_messages.json",
]

ARCHIVOS_BASE = [
    DATA_DIR / "facturas.xlsx",
    DATA_DIR / "historial_ejecuciones.xlsx",
]

PATRONES_AUDIT = [
    f"audit_detalle_{FECHA}.csv",
    f"audit_runs_{FECHA}.csv",
    f"audit_*_{FECHA}.csv",
]

PATRONES_LOGS = ["*.log", "*.txt"]

PALABRAS_SENSIBLES = (
    "SECRET",
    "PASSWORD",
    "PASS",
    "TOKEN",
    "KEY",
    "CLIENT_SECRET",
    "PRIVATE",
    "CERT",
)

NOMBRES_EXCLUIDOS = {
    ".env",
    ".env.local",
    ".env.production",
    "token_cache.json",
    "msal_cache.bin",
}


def rel(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(ROOT.resolve())).replace("\\", "/")
    except Exception:
        return str(path).replace("\\", "/")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def iso_mtime(path: Path) -> str:
    try:
        return datetime.datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds")
    except Exception:
        return ""


def acquire_lock() -> Tuple[bool, str]:
    LOCKS_DIR.mkdir(parents=True, exist_ok=True)
    now_ts = time.time()

    if LOCK_FILE.exists():
        age = now_ts - LOCK_FILE.stat().st_mtime
        if age < LOCK_TTL_SECONDS:
            return False, f"Lock activo: {LOCK_FILE} | edad={age:.0f}s | ttl={LOCK_TTL_SECONDS}s"
        try:
            LOCK_FILE.unlink()
        except Exception as exc:
            return False, f"No se pudo eliminar lock vencido: {LOCK_FILE} | {exc}"

    payload = {
        "script": "cierre_diario_seguro.py",
        "version": VERSION_CIERRE,
        "pid": os.getpid(),
        "started_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "host": socket.gethostname(),
    }
    LOCK_FILE.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return True, "Lock creado"


def release_lock() -> None:
    try:
        if LOCK_FILE.exists():
            LOCK_FILE.unlink()
    except Exception:
        pass


def es_sensible(linea: str) -> bool:
    if "=" not in linea:
        return False
    clave = linea.split("=", 1)[0].strip().upper()
    return any(palabra in clave for palabra in PALABRAS_SENSIBLES)


def crear_snapshot_env_redactado(destino_dir: Path) -> Optional[Dict[str, Any]]:
    env_path = ROOT / ".env"
    if not env_path.exists():
        return None

    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / f"snapshot_env_redactado_{FECHA}.txt"
    lineas_out: List[str] = []

    for linea in env_path.read_text(encoding="utf-8", errors="replace").splitlines():
        if not linea.strip() or linea.lstrip().startswith("#") or "=" not in linea:
            lineas_out.append(linea)
            continue
        clave, valor = linea.split("=", 1)
        if es_sensible(clave):
            lineas_out.append(f"{clave}=***REDACTADO***")
        else:
            lineas_out.append(f"{clave}={valor}")

    destino.write_text("\n".join(lineas_out) + "\n", encoding="utf-8")
    return info_archivo(destino, "config_redactada", ROOT / ".env")


def info_archivo(destino: Path, categoria: str, origen: Optional[Path] = None) -> Dict[str, Any]:
    st = destino.stat()
    return {
        "categoria": categoria,
        "origen": rel(origen) if origen else "generado",
        "destino": rel(destino),
        "nombre": destino.name,
        "bytes": st.st_size,
        "sha256": sha256_file(destino),
        "mtime_origen": iso_mtime(origen) if origen and origen.exists() else "",
        "mtime_destino": iso_mtime(destino),
    }


def copiar_archivo(origen: Path, subcarpeta: str, categoria: str) -> Optional[Dict[str, Any]]:
    if not origen.exists() or not origen.is_file():
        return None
    if origen.name in NOMBRES_EXCLUIDOS:
        return None

    destino_dir = CIERRE_DIA_DIR / subcarpeta
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / origen.name
    shutil.copy2(origen, destino)
    return info_archivo(destino, categoria, origen)


def archivos_unicos(rutas: List[Path]) -> List[Path]:
    vistos = set()
    out = []
    for p in rutas:
        try:
            key = str(p.resolve()).lower()
        except Exception:
            key = str(p).lower()
        if key not in vistos:
            vistos.add(key)
            out.append(p)
    return out


def recolectar_audits() -> List[Path]:
    rutas: List[Path] = []
    for base in (DATA_DIR, ROOT):
        if not base.exists():
            continue
        for patron in PATRONES_AUDIT:
            rutas.extend(base.glob(patron))
    return archivos_unicos([p for p in rutas if p.is_file()])


def recolectar_estado() -> List[Path]:
    rutas: List[Path] = []
    for nombre in NOMBRES_ESTADO:
        rutas.append(STATE_DIR / nombre)
    return [p for p in rutas if p.exists() and p.is_file()]


def recolectar_logs_recientes() -> List[Path]:
    if not LOGS_DIR.exists():
        return []
    limite = time.time() - (48 * 3600)
    rutas: List[Path] = []
    for patron in PATRONES_LOGS:
        rutas.extend(LOGS_DIR.glob(patron))
    return archivos_unicos([p for p in rutas if p.is_file() and p.stat().st_mtime >= limite])


def escribir_resumen(manifest: Dict[str, Any]) -> Dict[str, Any]:
    destino = CIERRE_DIA_DIR / f"RESUMEN_CIERRE_DIARIO_{FECHA}.txt"
    conteos: Dict[str, int] = {}
    for item in manifest["archivos"]:
        conteos[item["categoria"]] = conteos.get(item["categoria"], 0) + 1

    lineas = [
        "CIERRE DIARIO LOCAL SEGURO - FACTURAS JOYCO",
        "=" * 80,
        f"Version: {VERSION_CIERRE}",
        f"Fecha: {FECHA}",
        f"Generado: {manifest['generado_en']}",
        f"Root: {ROOT}",
        f"Carpeta cierre: {CIERRE_DIA_DIR}",
        "",
        "Contenido copiado:",
    ]
    for categoria, cantidad in sorted(conteos.items()):
        lineas.append(f"- {categoria}: {cantidad}")
    lineas.extend([
        "",
        f"Total archivos en manifest: {len(manifest['archivos'])}",
        f"Total bytes: {manifest['total_bytes']}",
        "",
        "Regla de seguridad:",
        "- No se copia .env con secretos reales.",
        "- Solo se genera snapshot .env redactado si existe archivo .env.",
        "- No se copian adjuntos, extraidos, ZIP/PDF/XML temporales ni carpetas pesadas.",
        "- Este cierre local debe subirse luego con subir_cierre_diario_sharepoint.py.",
        "",
    ])
    destino.write_text("\n".join(lineas), encoding="utf-8")
    return info_archivo(destino, "resumen", None)


def main() -> int:
    print("=" * 100)
    print("CIERRE DIARIO LOCAL SEGURO - FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_CIERRE}")
    print(f"Root: {ROOT}")
    print(f"Fecha: {FECHA}")
    print(f"Mes: {MES}")
    print(f"Carpeta local destino: {CIERRE_DIA_DIR}")
    print("-" * 100)

    ok_lock, msg_lock = acquire_lock()
    if not ok_lock:
        print(f"❌ {msg_lock}")
        print("⚠️ No se ejecuta el cierre para evitar cruces.")
        return 2

    try:
        CIERRE_DIA_DIR.mkdir(parents=True, exist_ok=True)

        manifest: Dict[str, Any] = {
            "tipo": "cierre_diario_local_seguro",
            "version": VERSION_CIERRE,
            "fecha": FECHA,
            "mes": MES,
            "generado_en": datetime.datetime.now().isoformat(timespec="seconds"),
            "root": str(ROOT),
            "host": socket.gethostname(),
            "platform": platform.platform(),
            "python": sys.version.replace("\n", " "),
            "archivos": [],
            "advertencias": [],
        }

        copiados = 0

        for origen in ARCHIVOS_BASE:
            info = copiar_archivo(origen, "01_excel", "excel_operativo")
            if info:
                manifest["archivos"].append(info)
                copiados += 1
            elif origen.name == "facturas.xlsx":
                manifest["advertencias"].append(f"No existe archivo obligatorio esperado: {rel(origen)}")

        for origen in recolectar_audits():
            info = copiar_archivo(origen, "02_auditoria", "auditoria")
            if info:
                manifest["archivos"].append(info)
                copiados += 1

        for origen in recolectar_estado():
            info = copiar_archivo(origen, "03_state", "state")
            if info:
                manifest["archivos"].append(info)
                copiados += 1

        for origen in recolectar_logs_recientes():
            info = copiar_archivo(origen, "04_logs_recientes", "logs_recientes")
            if info:
                manifest["archivos"].append(info)
                copiados += 1

        env_info = crear_snapshot_env_redactado(CIERRE_DIA_DIR / "05_config_redactada")
        if env_info:
            manifest["archivos"].append(env_info)
            copiados += 1

        manifest["total_archivos"] = len(manifest["archivos"])
        manifest["total_bytes"] = sum(int(item.get("bytes", 0) or 0) for item in manifest["archivos"])

        resumen_info = escribir_resumen(manifest)
        manifest["archivos"].append(resumen_info)
        manifest["total_archivos"] = len(manifest["archivos"])
        manifest["total_bytes"] = sum(int(item.get("bytes", 0) or 0) for item in manifest["archivos"])

        manifest_path = CIERRE_DIA_DIR / f"manifest_cierre_diario_{FECHA}.json"
        manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")

        print(f"✅ Archivos copiados: {copiados}")
        print(f"✅ Manifest: {manifest_path}")
        print(f"✅ Resumen: {CIERRE_DIA_DIR / f'RESUMEN_CIERRE_DIARIO_{FECHA}.txt'}")
        print(f"✅ Total archivos manifest: {manifest['total_archivos']}")
        print(f"✅ Total bytes: {manifest['total_bytes']}")

        if manifest["advertencias"]:
            print("⚠️ Advertencias:")
            for adv in manifest["advertencias"]:
                print(f"  - {adv}")

        if copiados <= 0:
            print("❌ No se copió ningún archivo operativo. Revisa rutas de data/.")
            return 1

        print("=" * 100)
        print("✅ Cierre diario local generado correctamente.")
        print("Siguiente paso: python scripts\\subir_cierre_diario_sharepoint.py")
        print("=" * 100)
        return 0

    except Exception as exc:
        print(f"❌ Error generando cierre diario local: {exc}")
        print("⚠️ No subas ni archives nada hasta revisar el error.")
        return 1
    finally:
        release_lock()


if __name__ == "__main__":
    raise SystemExit(main())
