# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Cierre trimestral de facturas

Modos:
- --dry-run: solo diagnostica, no modifica nada.
- --local: genera cierre trimestral local de prueba, sin reemplazar facturas.xlsx y sin tocar SharePoint.

Estructura generada:
data/cierres_trimestrales/YYYY/T#
├── 01_Excel_Cierre
└── 02_Soportes_Tecnicos
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
STATE_PATH = DATA_DIR / "state" / "cierre_trimestral_state.json"
FACTURAS_PATH = DATA_DIR / "facturas.xlsx"
BACKUPS_MENSUALES_DIR = DATA_DIR / "backups_mensuales"
CIERRES_TRIMESTRALES_DIR = DATA_DIR / "cierres_trimestrales"

VERSION = "2026-06-18-CIERRE-TRIMESTRAL-LOCAL-V3-SIN-PREVIEW"

HEADERS_ESPERADOS = [
    "Radicado",
    "ProyectoProceso",
    "Archivo",
    "Empresa emisora",
    "CUFE",
    "Ciudad emisora",
    "Código ciudad",
    "NIT",
    "Cliente",
    "Número de factura",
    "Año",
    "Mes",
    "Día",
    "Tipo de contribuyente",
    "Actividad económica",
    "DESCRIPCIÓN",
    "Concepto",
    "VALOR",
    "Estado_calidad",
]


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def cargar_estado() -> dict:
    if not STATE_PATH.exists():
        raise RuntimeError(
            f"No existe el estado trimestral: {STATE_PATH}\n"
            "No se puede cerrar trimestre. Primero inicializa data/state/cierre_trimestral_state.json."
        )

    with STATE_PATH.open("r", encoding="utf-8-sig") as f:
        estado = json.load(f)

    requeridos = [
        "periodo_activo",
        "fecha_inicio_periodo_activo",
        "proximo_cierre_estimado",
        "estado",
    ]

    faltantes = [k for k in requeridos if not estado.get(k)]
    if faltantes:
        raise RuntimeError(f"Estado trimestral incompleto. Faltan campos: {faltantes}")

    if str(estado.get("estado")).upper() != "ACTIVO":
        raise RuntimeError(f"Estado trimestral no está ACTIVO: {estado.get('estado')}")

    return estado


def validar_fecha(valor: str, campo: str) -> datetime:
    try:
        return datetime.strptime(valor, "%Y-%m-%d")
    except Exception as exc:
        raise RuntimeError(f"Fecha inválida en {campo}: {valor}. Formato esperado YYYY-MM-DD.") from exc


def diagnosticar_excel() -> dict:
    if not FACTURAS_PATH.exists():
        raise RuntimeError(f"No existe el Excel principal: {FACTURAS_PATH}")

    wb = load_workbook(FACTURAS_PATH, read_only=False, data_only=False)

    if "Facturas" not in wb.sheetnames:
        wb.close()
        raise RuntimeError("El Excel no tiene la hoja requerida: Facturas")

    ws = wb["Facturas"]

    headers = [cell.value for cell in ws[1]]
    if headers != HEADERS_ESPERADOS:
        wb.close()
        raise RuntimeError(
            "Los encabezados del Excel no coinciden con la estructura esperada.\n"
            f"Detectados: {headers}\n"
            f"Esperados: {HEADERS_ESPERADOS}"
        )

    tablas = list(ws.tables.keys())
    tabla_ref = None
    if "TblFacturas" in ws.tables:
        tabla_ref = ws.tables["TblFacturas"].ref

    info = {
        "hojas": wb.sheetnames,
        "hoja_principal": ws.title,
        "filas": ws.max_row,
        "columnas": ws.max_column,
        "filas_datos": max(ws.max_row - 1, 0),
        "tablas": tablas,
        "tbl_facturas_ref": tabla_ref,
    }

    wb.close()
    return info


def buscar_ultimo_backup_mensual() -> Optional[Path]:
    if not BACKUPS_MENSUALES_DIR.exists():
        return None

    backups = sorted(
        BACKUPS_MENSUALES_DIR.rglob("backup_mensual_*.zip"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    return backups[0] if backups else None


def datos_periodo(estado: dict) -> Dict[str, str]:
    periodo = estado["periodo_activo"]
    fecha_inicio = estado["fecha_inicio_periodo_activo"]
    fecha_fin = estado["proximo_cierre_estimado"]

    validar_fecha(fecha_inicio, "fecha_inicio_periodo_activo")
    validar_fecha(fecha_fin, "proximo_cierre_estimado")

    if "-" in periodo:
        anio, trimestre = periodo.split("-", 1)
    else:
        anio = fecha_inicio[:4]
        trimestre = periodo

    nombre_archivo = f"facturas_{fecha_inicio}_a_{fecha_fin}.xlsx"

    carpeta_base = CIERRES_TRIMESTRALES_DIR / anio / trimestre
    carpeta_excel = carpeta_base / "01_Excel_Cierre"
    carpeta_soportes = carpeta_base / "02_Soportes_Tecnicos"
    destino_local = carpeta_excel / nombre_archivo

    return {
        "periodo": periodo,
        "anio": anio,
        "trimestre": trimestre,
        "fecha_inicio": fecha_inicio,
        "fecha_fin": fecha_fin,
        "nombre_archivo": nombre_archivo,
        "carpeta_base": str(carpeta_base),
        "carpeta_excel": str(carpeta_excel),
        "carpeta_soportes": str(carpeta_soportes),
        "destino_local": str(destino_local),
    }


def escribir_manifest(
    manifest_path: Path,
    periodo: Dict[str, str],
    info_excel: dict,
    backup: Path,
    archivo_cerrado: Path,
    archivo_cerrado_reutilizado: bool,
) -> dict:
    manifest = {
        "tipo": "cierre_trimestral_local",
        "version_script": VERSION,
        "generado_en": datetime.now().isoformat(timespec="seconds"),
        "root": str(ROOT),
        "periodo": periodo["periodo"],
        "fecha_inicio": periodo["fecha_inicio"],
        "fecha_fin": periodo["fecha_fin"],
        "excel_original": str(FACTURAS_PATH),
        "excel_original_bytes": FACTURAS_PATH.stat().st_size,
        "excel_original_sha256": sha256_file(FACTURAS_PATH),
        "excel_cerrado": str(archivo_cerrado),
        "excel_cerrado_bytes": archivo_cerrado.stat().st_size,
        "excel_cerrado_sha256": sha256_file(archivo_cerrado),
        "excel_cerrado_reutilizado": archivo_cerrado_reutilizado,
        "backup_mensual_usado": str(backup),
        "backup_mensual_bytes": backup.stat().st_size,
        "backup_mensual_sha256": sha256_file(backup),
        "excel_info": info_excel,
        "carpeta_excel_cierre": periodo["carpeta_excel"],
        "carpeta_soportes_tecnicos": periodo["carpeta_soportes"],
        "nota": (
            "Modo local. No reemplaza data/facturas.xlsx, "
            "no toca SharePoint y no actualiza cierre_trimestral_state.json."
        ),
    }

    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def escribir_resumen(resumen_path: Path, manifest: dict) -> None:
    lines = [
        "CIERRE TRIMESTRAL FACTURAS - LOCAL",
        "=" * 80,
        f"Version: {manifest['version_script']}",
        f"Generado en: {manifest['generado_en']}",
        f"Periodo: {manifest['periodo']}",
        f"Fecha inicio: {manifest['fecha_inicio']}",
        f"Fecha fin: {manifest['fecha_fin']}",
        "",
        "Carpeta para contabilidad:",
        f"- {manifest['carpeta_excel_cierre']}",
        "",
        "Carpeta de soportes tecnicos:",
        f"- {manifest['carpeta_soportes_tecnicos']}",
        "",
        "Archivo original:",
        f"- {manifest['excel_original']}",
        f"- SHA256: {manifest['excel_original_sha256']}",
        "",
        "Archivo cerrado generado/reutilizado:",
        f"- {manifest['excel_cerrado']}",
        f"- SHA256: {manifest['excel_cerrado_sha256']}",
        f"- Reutilizado: {manifest['excel_cerrado_reutilizado']}",
        "",
        "Backup mensual usado como respaldo previo:",
        f"- {manifest['backup_mensual_usado']}",
        f"- SHA256: {manifest['backup_mensual_sha256']}",
        "",
        "Importante:",
        "- Este modo NO reemplaza data/facturas.xlsx.",
        "- Este modo NO toca SharePoint.",
        "- Este modo NO actualiza cierre_trimestral_state.json.",
        "- No genera facturas_LIMPIO_PREVIEW para no ocupar espacio innecesario.",
        "=" * 80,
    ]
    resumen_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def imprimir_plan(estado: dict, info_excel: dict, backup: Optional[Path]) -> Dict[str, str]:
    periodo = datos_periodo(estado)

    print("✅ Estado trimestral cargado correctamente.")
    print(f"Periodo activo: {periodo['periodo']}")
    print(f"Fecha inicio periodo: {periodo['fecha_inicio']}")
    print(f"Fecha cierre estimada: {periodo['fecha_fin']}")

    print("-" * 100)
    print("✅ Excel principal validado.")
    print(f"Archivo: {FACTURAS_PATH}")
    print(f"Hojas: {info_excel['hojas']}")
    print(f"Hoja principal: {info_excel['hoja_principal']}")
    print(f"Filas totales: {info_excel['filas']}")
    print(f"Filas de datos: {info_excel['filas_datos']}")
    print(f"Columnas: {info_excel['columnas']}")
    print(f"Tablas: {info_excel['tablas']}")
    print(f"TblFacturas ref: {info_excel['tbl_facturas_ref']}")

    print("-" * 100)
    if backup:
        print("✅ Backup mensual disponible.")
        print(f"Último backup: {backup}")
        print(f"Bytes: {backup.stat().st_size}")
    else:
        print("⚠️ No se encontró backup mensual ZIP.")
        print("En modo local/real, el cierre trimestral debe bloquearse si no hay backup previo.")

    print("-" * 100)
    print("PLAN DE CIERRE TRIMESTRAL:")
    print("1. Copiar/reutilizar Excel actual en carpeta de contabilidad:")
    print(f"   {periodo['destino_local']}")
    print("2. Guardar manifest y resumen en carpeta técnica:")
    print(f"   {periodo['carpeta_soportes']}")
    print("3. No generar Excel LIMPIO_PREVIEW en este modo.")
    print("4. En una fase posterior: crear nuevo data/facturas.xlsx limpio.")
    print("5. En una fase posterior: subir a SharePoint y reemplazar activo.")
    print("6. En una fase posterior: actualizar cierre_trimestral_state.json.")

    return periodo


def preparar_archivo_cerrado(carpeta_excel: Path, archivo_destino: Path) -> tuple[Path, bool]:
    """
    Si el archivo cerrado ya existe y coincide con el Excel actual, se reutiliza.
    Si existe pero es diferente, se crea una copia con timestamp.
    Si no existe, se crea normalmente.
    """
    if not archivo_destino.exists():
        shutil.copy2(FACTURAS_PATH, archivo_destino)
        return archivo_destino, False

    hash_original = sha256_file(FACTURAS_PATH)
    hash_existente = sha256_file(archivo_destino)

    if hash_original == hash_existente:
        return archivo_destino, True

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nuevo_destino = carpeta_excel / (
        f"facturas_{archivo_destino.stem.replace('facturas_', '')}_{stamp}.xlsx"
    )
    shutil.copy2(FACTURAS_PATH, nuevo_destino)
    return nuevo_destino, False


def ejecutar_local(periodo: Dict[str, str], info_excel: dict, backup: Optional[Path]) -> int:
    if backup is None:
        raise RuntimeError("No hay backup mensual previo. Se bloquea el cierre local.")

    carpeta_excel = Path(periodo["carpeta_excel"])
    carpeta_soportes = Path(periodo["carpeta_soportes"])
    archivo_destino = Path(periodo["destino_local"])
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    carpeta_excel.mkdir(parents=True, exist_ok=True)
    carpeta_soportes.mkdir(parents=True, exist_ok=True)

    archivo_cerrado, reutilizado = preparar_archivo_cerrado(carpeta_excel, archivo_destino)

    manifest_path = carpeta_soportes / f"manifest_cierre_trimestral_{periodo['periodo']}_{stamp}.json"
    resumen_path = carpeta_soportes / f"RESUMEN_CIERRE_TRIMESTRAL_{periodo['periodo']}_{stamp}.txt"

    manifest = escribir_manifest(
        manifest_path=manifest_path,
        periodo=periodo,
        info_excel=info_excel,
        backup=backup,
        archivo_cerrado=archivo_cerrado,
        archivo_cerrado_reutilizado=reutilizado,
    )
    escribir_resumen(resumen_path, manifest)

    print("-" * 100)
    print("✅ CIERRE TRIMESTRAL LOCAL GENERADO.")
    print(f"Excel cierre contabilidad: {archivo_cerrado}")
    print(f"Manifest: {manifest_path}")
    print(f"Resumen: {resumen_path}")
    print("-" * 100)
    print("✅ No se generó facturas_LIMPIO_PREVIEW.")
    print("✅ No se reemplazó data/facturas.xlsx.")
    print("✅ No se tocó SharePoint.")
    print("✅ No se actualizó cierre_trimestral_state.json.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Solo diagnostica, no modifica archivos.")
    parser.add_argument("--local", action="store_true", help="Genera cierre local, sin reemplazar original.")
    args = parser.parse_args()

    modo = "LOCAL" if args.local else "DRY RUN"

    print("=" * 100)
    print(f"CIERRE TRIMESTRAL FACTURAS - {modo}")
    print("=" * 100)
    print(f"Versión: {VERSION}")
    print(f"Root: {ROOT}")
    print("-" * 100)

    try:
        estado = cargar_estado()
        info_excel = diagnosticar_excel()
        backup = buscar_ultimo_backup_mensual()
        periodo = imprimir_plan(estado, info_excel, backup)

        if args.local:
            ejecutar_local(periodo, info_excel, backup)
        else:
            print("-" * 100)
            print("✅ DRY RUN finalizado. No se modificó ningún archivo.")

        print("=" * 100)
        return 0

    except Exception as exc:
        print(f"❌ Error en cierre trimestral: {exc}")
        print("No se debe continuar hasta revisar el error.")
        print("=" * 100)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
