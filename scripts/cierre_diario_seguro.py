import os
import json
import shutil
import hashlib
import datetime
from pathlib import Path

VERSION_CIERRE = "2026-06-11-CIERRE-DIARIO-SEGURO-V2-SNAPSHOT-SEGURO"

# Importante:
# No usamos Path.cwd() como raíz principal, porque cuando esto quede
# programado en Windows, el "Start in" puede cambiar. La raíz se calcula
# desde la ubicación real del script: scripts/cierre_diario_seguro.py
ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
AUDIT_DIR = DATA_DIR / "audit"
STATE_DIR = DATA_DIR / "state"

FECHA = datetime.datetime.now().strftime("%Y-%m-%d")
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

DESTINO = DATA_DIR / "cierres_diarios" / FECHA

ARCHIVOS_A_RESPALDAR = [
    AUDIT_DIR / f"audit_runs_{FECHA}.csv",
    AUDIT_DIR / f"audit_detalle_{FECHA}.csv",
    STATE_DIR / "processed_messages.json",
    STATE_DIR / "attachment_index_store.json",
]

LOCKS_POSIBLES = [
    DATA_DIR / "locks" / "aprobadas.lock",
    DATA_DIR / "aprobadas.lock",
    ROOT / "aprobadas.lock",
]

ENV_PATH = ROOT / ".env"

KEYWORDS_SENSIBLES = (
    "SECRET",
    "PASSWORD",
    "PASS",
    "TOKEN",
    "KEY",
    "PRIVATE",
    "CREDENTIAL",
    "CLIENT_SECRET",
    "TENANT",
    "CLIENT_ID",
    "AUTHORITY",
    "GRAPH",
)

def size_safe(path: Path) -> int:
    try:
        return path.stat().st_size
    except Exception:
        return 0

def sha256_file(path: Path) -> str:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return ""

def existe_lock_activo() -> list[str]:
    activos = []
    for lock in LOCKS_POSIBLES:
        if lock.exists():
            activos.append(str(lock))
    return activos

def copiar_archivo(src: Path, destino_dir: Path) -> dict:
    info = {
        "origen": str(src),
        "existe": src.exists(),
        "copiado": False,
        "destino": "",
        "size_bytes": 0,
        "sha256": "",
        "error": "",
    }

    if not src.exists():
        return info

    try:
        destino_dir.mkdir(parents=True, exist_ok=True)
        dst = destino_dir / src.name
        shutil.copy2(src, dst)

        info["copiado"] = True
        info["destino"] = str(dst)
        info["size_bytes"] = size_safe(dst)
        info["sha256"] = sha256_file(dst)
        return info
    except Exception as e:
        info["error"] = str(e)
        return info

def es_clave_sensible(key: str) -> bool:
    k = str(key or "").upper().strip()
    if not k:
        return True
    return any(word in k for word in KEYWORDS_SENSIBLES)

def redacted_value(key: str, value: str) -> str:
    if es_clave_sensible(key):
        if value:
            return "***REDACTED***"
        return ""

    # No guardar valores larguísimos aunque la clave parezca segura.
    if value and len(value) > 300:
        return "***REDACTED_LONG_VALUE***"

    return value

def generar_env_snapshot_seguro(destino_dir: Path) -> dict:
    """
    Genera snapshot seguro del .env sin copiar secretos reales.
    El objetivo es dejar trazabilidad de configuración operativa,
    no respaldo de credenciales.
    """
    info = {
        "origen": str(ENV_PATH),
        "existe": ENV_PATH.exists(),
        "copiado": False,
        "destino": "",
        "size_bytes": 0,
        "sha256": "",
        "error": "",
        "tipo": "env_snapshot_redacted",
    }

    if not ENV_PATH.exists():
        return info

    try:
        destino_dir.mkdir(parents=True, exist_ok=True)
        dst = destino_dir / f"env_snapshot_redacted_{STAMP}.txt"

        lineas_out = []
        lineas_out.append("# Snapshot seguro de .env")
        lineas_out.append(f"# Generado: {datetime.datetime.now().isoformat(timespec='seconds')}")
        lineas_out.append("# Los valores sensibles se reemplazan por ***REDACTED***")
        lineas_out.append("")

        with open(ENV_PATH, "r", encoding="utf-8", errors="replace") as f:
            for raw in f.readlines():
                line = raw.rstrip("\n")

                if not line.strip():
                    lineas_out.append("")
                    continue

                if line.strip().startswith("#"):
                    lineas_out.append(line)
                    continue

                if "=" not in line:
                    lineas_out.append("# LINEA_NO_KEY_VALUE_REDACTED")
                    continue

                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()

                safe_value = redacted_value(key, value)
                lineas_out.append(f"{key}={safe_value}")

        dst.write_text("\n".join(lineas_out) + "\n", encoding="utf-8")

        info["copiado"] = True
        info["destino"] = str(dst)
        info["size_bytes"] = size_safe(dst)
        info["sha256"] = sha256_file(dst)
        return info

    except Exception as e:
        info["error"] = str(e)
        return info

def main():
    print("=" * 100)
    print("CIERRE DIARIO SEGURO - AUTOMATIZACIÓN FACTURAS JOYCO")
    print("=" * 100)
    print(f"Versión: {VERSION_CIERRE}")
    print(f"Root: {ROOT}")
    print(f"Fecha: {FECHA}")
    print(f"Destino: {DESTINO}")
    print("-" * 100)

    locks = existe_lock_activo()
    if locks:
        print("⚠️ Hay lock activo. No se hace cierre para evitar copiar estado durante ejecución.")
        for l in locks:
            print(f"  - {l}")

        DESTINO.mkdir(parents=True, exist_ok=True)
        manifest = {
            "version": VERSION_CIERRE,
            "fecha": FECHA,
            "timestamp": STAMP,
            "estado": "BLOQUEADO_POR_LOCK",
            "root": str(ROOT),
            "locks": locks,
            "archivos": [],
        }

        manifest_path = DESTINO / f"manifest_cierre_{STAMP}.json"
        manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")

        print(f"Manifest generado: {manifest_path}")
        print("=" * 100)
        return 2

    DESTINO.mkdir(parents=True, exist_ok=True)

    resultados = []

    for src in ARCHIVOS_A_RESPALDAR:
        r = copiar_archivo(src, DESTINO)
        resultados.append(r)

        if r["copiado"]:
            print(f"✅ Copiado: {src.name} ({r['size_bytes']} bytes)")
        elif r["existe"]:
            print(f"❌ Error copiando: {src.name} | {r['error']}")
        else:
            print(f"ℹ️ No existe, se omite: {src}")

    env_info = generar_env_snapshot_seguro(DESTINO)
    resultados.append(env_info)

    if env_info["copiado"]:
        print(f"✅ Snapshot seguro .env: {Path(env_info['destino']).name} ({env_info['size_bytes']} bytes)")
    elif env_info["existe"]:
        print(f"❌ Error generando snapshot seguro .env: {env_info['error']}")
    else:
        print("ℹ️ No existe .env, se omite snapshot seguro.")

    errores = [r for r in resultados if r["existe"] and not r["copiado"]]

    manifest = {
        "version": VERSION_CIERRE,
        "fecha": FECHA,
        "timestamp": STAMP,
        "estado": "OK" if not errores else "ERROR",
        "root": str(ROOT),
        "destino": str(DESTINO),
        "locks": [],
        "archivos": resultados,
    }

    manifest_path = DESTINO / f"manifest_cierre_{STAMP}.json"
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")

    print("-" * 100)
    print(f"Manifest generado: {manifest_path}")

    if errores:
        print("❌ Cierre terminado con errores.")
        print("=" * 100)
        return 1

    print("✅ Cierre diario seguro terminado correctamente.")
    print("=" * 100)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
