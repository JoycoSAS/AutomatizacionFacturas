# -*- coding: utf-8 -*-
"""
JOYCO - Facturas Procesador
Cierre trimestral de facturas

Modos:
- --dry-run: solo diagnostica, no modifica nada.
- --local: genera cierre trimestral local de prueba, sin reemplazar facturas.xlsx y sin tocar SharePoint.
- --real --confirmar CERRAR_TRIMESTRE: cierre real local/VPS, reemplaza data/facturas.xlsx y actualiza state.

Estructura generada:
data/cierres_trimestrales/YYYY/T#
├── 01_Excel_Cierre
└── 02_Soportes_Tecnicos
"""

from __future__ import annotations

import argparse
import calendar
import hashlib
import json
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
STATE_PATH = DATA_DIR / "state" / "cierre_trimestral_state.json"
FACTURAS_PATH = DATA_DIR / "facturas.xlsx"
BACKUPS_MENSUALES_DIR = DATA_DIR / "backups_mensuales"
CIERRES_TRIMESTRALES_DIR = DATA_DIR / "cierres_trimestrales"

VERSION = "2026-06-18-CIERRE-TRIMESTRAL-REAL-V4-LOCAL-VPS"

CONFIRMACION_REAL = "CERRAR_TRIMESTRE"

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


def guardar_estado_atomico(estado: dict) -> None:
    tmp = STATE_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(STATE_PATH)


def parse_fecha(valor: str, campo: str) -> date:
    try:
        return datetime.strptime(valor, "%Y-%m-%d").date()
    except Exception as exc:
        raise RuntimeError(f"Fecha inválida en {campo}: {valor}. Formato esperado YYYY-MM-DD.") from exc


def add_months(fecha: date, meses: int) -> date:
    mes_total = fecha.month - 1 + meses
    anio = fecha.year + mes_total // 12
    mes = mes_total % 12 + 1
    dia = min(fecha.day, calendar.monthrange(anio, mes)[1])
    return date(anio, mes, dia)


def calcular_siguiente_periodo(fecha_fin_actual: date) -> Dict[str, str]:
    siguiente_inicio = fecha_fin_actual + timedelta(days=1)
    siguiente_fin = add_months(siguiente_inicio, 3) - timedelta(days=1)
    trimestre = ((siguiente_inicio.month - 1) // 3) + 1
    periodo = f"{siguiente_inicio.year}-T{trimestre}"

    return {
        "periodo_activo": periodo,
        "fecha_inicio_periodo_activo": siguiente_inicio.isoformat(),
        "proximo_cierre_estimado": siguiente_fin.isoformat(),
    }


def diagnosticar_excel(path: Path = FACTURAS_PATH) -> dict:
    if not path.exists():
        raise RuntimeError(f"No existe el Excel: {path}")

    wb = load_workbook(path, read_only=False, data_only=False)

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
        "archivo": str(path),
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

    parse_fecha(fecha_inicio, "fecha_inicio_periodo_activo")
    parse_fecha(fecha_fin, "proximo_cierre_estimado")

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


def crear_excel_limpio(destino: Path) -> None:
    """
    Crea un Excel limpio con la misma estructura base:
    - Hoja Facturas.
    - Encabezados.
    - Tabla TblFacturas A1:S1.
    """
    wb = load_workbook(FACTURAS_PATH, read_only=False, data_only=False)
    ws = wb["Facturas"]

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    ws.tables.clear()

    tab = Table(displayName="TblFacturas", ref="A1:S1")
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
    wb.close()


def preparar_archivo_cerrado(carpeta_excel: Path, archivo_destino: Path) -> Tuple[Path, bool]:
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


def escribir_manifest(
    manifest_path: Path,
    periodo: Dict[str, str],
    info_excel: dict,
    backup: Path,
    archivo_cerrado: Path,
    archivo_cerrado_reutilizado: bool,
    modo: str,
    extras: Optional[dict] = None,
) -> dict:
    manifest = {
        "tipo": modo,
        "version_script": VERSION,
        "generado_en": datetime.now().isoformat(timespec="seconds"),
        "root": str(ROOT),
        "periodo": periodo["periodo"],
        "fecha_inicio": periodo["fecha_inicio"],
        "fecha_fin": periodo["fecha_fin"],
        "excel_original": str(FACTURAS_PATH),
        "excel_original_existe": FACTURAS_PATH.exists(),
        "excel_cerrado": str(archivo_cerrado),
        "excel_cerrado_bytes": archivo_cerrado.stat().st_size,
        "excel_cerrado_sha256": sha256_file(archivo_cerrado),
        "excel_cerrado_reutilizado": archivo_cerrado_reutilizado,
        "backup_mensual_usado": str(backup),
        "backup_mensual_bytes": backup.stat().st_size,
        "backup_mensual_sha256": sha256_file(backup),
        "excel_info_antes_cierre": info_excel,
        "carpeta_excel_cierre": periodo["carpeta_excel"],
        "carpeta_soportes_tecnicos": periodo["carpeta_soportes"],
    }

    if FACTURAS_PATH.exists():
        manifest["excel_original_bytes"] = FACTURAS_PATH.stat().st_size
        manifest["excel_original_sha256"] = sha256_file(FACTURAS_PATH)

    if extras:
        manifest.update(extras)

    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def escribir_resumen(resumen_path: Path, manifest: dict) -> None:
    lines = [
        f"CIERRE TRIMESTRAL FACTURAS - {manifest['tipo']}",
        "=" * 80,
        f"Version: {manifest['version_script']}",
        f"Generado en: {manifest['generado_en']}",
        f"Periodo cerrado: {manifest['periodo']}",
        f"Fecha inicio: {manifest['fecha_inicio']}",
        f"Fecha fin: {manifest['fecha_fin']}",
        "",
        "Carpeta para contabilidad:",
        f"- {manifest['carpeta_excel_cierre']}",
        "",
        "Carpeta de soportes tecnicos:",
        f"- {manifest['carpeta_soportes_tecnicos']}",
        "",
        "Archivo cerrado:",
        f"- {manifest['excel_cerrado']}",
        f"- SHA256: {manifest['excel_cerrado_sha256']}",
        f"- Reutilizado: {manifest['excel_cerrado_reutilizado']}",
        "",
        "Backup mensual usado como respaldo previo:",
        f"- {manifest['backup_mensual_usado']}",
        f"- SHA256: {manifest['backup_mensual_sha256']}",
    ]

    if manifest.get("nuevo_excel_activo"):
        lines += [
            "",
            "Nuevo Excel activo:",
            f"- {manifest['nuevo_excel_activo']}",
            f"- SHA256: {manifest.get('nuevo_excel_activo_sha256')}",
            f"- Filas: {manifest.get('nuevo_excel_activo_filas')}",
            f"- Tabla: {manifest.get('nuevo_excel_activo_tbl_facturas_ref')}",
        ]

    if manifest.get("nuevo_periodo_activo"):
        lines += [
            "",
            "Nuevo periodo activo:",
            f"- {manifest['nuevo_periodo_activo']}",
            f"- Inicio: {manifest['nuevo_fecha_inicio_periodo_activo']}",
            f"- Proximo cierre: {manifest['nuevo_proximo_cierre_estimado']}",
        ]

    lines += [
        "",
        "Importante:",
        "- En modo LOCAL no se reemplaza data/facturas.xlsx.",
        "- En modo REAL sí se reemplaza data/facturas.xlsx por un Excel limpio.",
        "- Este script no toca SharePoint todavía.",
        "=" * 80,
    ]

    resumen_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def imprimir_plan(estado: dict, info_excel: dict, backup: Optional[Path], modo: str) -> Dict[str, str]:
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
    print("3. No generar Excel LIMPIO_PREVIEW.")

    if modo == "REAL":
        print("4. Crear nuevo data/facturas.xlsx limpio.")
        print("5. Reemplazar Excel activo local.")
        print("6. Actualizar cierre_trimestral_state.json al siguiente trimestre.")
        print("7. SharePoint queda para fase posterior.")
    else:
        print("4. No reemplazar data/facturas.xlsx.")
        print("5. No actualizar cierre_trimestral_state.json.")

    return periodo


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
        modo="LOCAL",
        extras={
            "nota": (
                "Modo local. No reemplaza data/facturas.xlsx, "
                "no toca SharePoint y no actualiza cierre_trimestral_state.json."
            )
        },
    )
    escribir_resumen(resumen_path, manifest)

    print("-" * 100)
    print("✅ CIERRE TRIMESTRAL LOCAL GENERADO.")
    print(f"Excel cierre contabilidad: {archivo_cerrado}")
    print(f"Manifest: {manifest_path}")
    print(f"Resumen: {resumen_path}")
    print("-" * 100)
    print("✅ No se reemplazó data/facturas.xlsx.")
    print("✅ No se tocó SharePoint.")
    print("✅ No se actualizó cierre_trimestral_state.json.")
    return 0


def validar_ejecucion_real(estado: dict, confirmar: Optional[str]) -> date:
    if confirmar != CONFIRMACION_REAL:
        raise RuntimeError(
            "Cierre real bloqueado. Para ejecutarlo usa:\n"
            f"python scripts\\cierre_trimestral_facturas.py --real --confirmar {CONFIRMACION_REAL}"
        )

    fecha_fin = parse_fecha(estado["proximo_cierre_estimado"], "proximo_cierre_estimado")
    hoy = date.today()

    if hoy < fecha_fin:
        raise RuntimeError(
            "Cierre real bloqueado por fecha.\n"
            f"Periodo activo: {estado['periodo_activo']}\n"
            f"Fecha de cierre estimada: {fecha_fin.isoformat()}\n"
            f"Fecha actual: {hoy.isoformat()}\n"
            "No se permite cerrar antes de que se cumpla el trimestre."
        )

    return hoy


def ejecutar_real(periodo: Dict[str, str], info_excel: dict, backup: Optional[Path], estado: dict) -> int:
    if backup is None:
        raise RuntimeError("No hay backup mensual previo. Se bloquea el cierre real.")

    carpeta_excel = Path(periodo["carpeta_excel"])
    carpeta_soportes = Path(periodo["carpeta_soportes"])
    archivo_destino = Path(periodo["destino_local"])
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    carpeta_excel.mkdir(parents=True, exist_ok=True)
    carpeta_soportes.mkdir(parents=True, exist_ok=True)

    archivo_cerrado, reutilizado = preparar_archivo_cerrado(carpeta_excel, archivo_destino)

    # Copia adicional de seguridad del Excel activo antes de reemplazar.
    respaldo_activo = carpeta_soportes / f"facturas_ACTIVO_ANTES_REEMPLAZO_{periodo['periodo']}_{stamp}.xlsx"
    shutil.copy2(FACTURAS_PATH, respaldo_activo)

    # Backup del state antes de actualizarlo.
    respaldo_state = carpeta_soportes / f"cierre_trimestral_state_ANTES_{periodo['periodo']}_{stamp}.json"
    shutil.copy2(STATE_PATH, respaldo_state)

    # Crear Excel limpio temporal y validarlo antes de reemplazar.
    nuevo_tmp = DATA_DIR / f"facturas_NUEVO_TMP_{periodo['periodo']}_{stamp}.xlsx"
    crear_excel_limpio(nuevo_tmp)
    info_nuevo = diagnosticar_excel(nuevo_tmp)

    if info_nuevo["filas"] != 1 or info_nuevo["columnas"] != 19:
        raise RuntimeError(f"Excel limpio temporal inválido: {info_nuevo}")

    if info_nuevo["tbl_facturas_ref"] != "A1:S1":
        raise RuntimeError(f"Tabla del Excel limpio temporal inválida: {info_nuevo}")

    # Reemplazo local controlado.
    FACTURAS_PATH.unlink()
    shutil.move(str(nuevo_tmp), str(FACTURAS_PATH))

    info_activo_nuevo = diagnosticar_excel(FACTURAS_PATH)

    fecha_fin = parse_fecha(periodo["fecha_fin"], "fecha_fin")
    siguiente = calcular_siguiente_periodo(fecha_fin)

    nuevo_estado = dict(estado)
    nuevo_estado["ultimo_cierre_trimestral"] = periodo["fecha_fin"]
    nuevo_estado["ultimo_archivo_generado"] = str(archivo_cerrado)
    nuevo_estado["periodo_activo"] = siguiente["periodo_activo"]
    nuevo_estado["fecha_inicio_periodo_activo"] = siguiente["fecha_inicio_periodo_activo"]
    nuevo_estado["proximo_cierre_estimado"] = siguiente["proximo_cierre_estimado"]
    nuevo_estado["estado"] = "ACTIVO"
    nuevo_estado["actualizado_en"] = datetime.now().isoformat(timespec="seconds")
    nuevo_estado["version"] = "2026-06-18-CIERRE-TRIMESTRAL-STATE-V2-POST-CIERRE"

    guardar_estado_atomico(nuevo_estado)

    manifest_path = carpeta_soportes / f"manifest_cierre_trimestral_REAL_{periodo['periodo']}_{stamp}.json"
    resumen_path = carpeta_soportes / f"RESUMEN_CIERRE_TRIMESTRAL_REAL_{periodo['periodo']}_{stamp}.txt"

    manifest = escribir_manifest(
        manifest_path=manifest_path,
        periodo=periodo,
        info_excel=info_excel,
        backup=backup,
        archivo_cerrado=archivo_cerrado,
        archivo_cerrado_reutilizado=reutilizado,
        modo="REAL",
        extras={
            "respaldo_excel_activo_antes_reemplazo": str(respaldo_activo),
            "respaldo_excel_activo_antes_reemplazo_sha256": sha256_file(respaldo_activo),
            "respaldo_state_antes_actualizar": str(respaldo_state),
            "nuevo_excel_activo": str(FACTURAS_PATH),
            "nuevo_excel_activo_bytes": FACTURAS_PATH.stat().st_size,
            "nuevo_excel_activo_sha256": sha256_file(FACTURAS_PATH),
            "nuevo_excel_activo_filas": info_activo_nuevo["filas"],
            "nuevo_excel_activo_columnas": info_activo_nuevo["columnas"],
            "nuevo_excel_activo_tbl_facturas_ref": info_activo_nuevo["tbl_facturas_ref"],
            "nuevo_periodo_activo": nuevo_estado["periodo_activo"],
            "nuevo_fecha_inicio_periodo_activo": nuevo_estado["fecha_inicio_periodo_activo"],
            "nuevo_proximo_cierre_estimado": nuevo_estado["proximo_cierre_estimado"],
            "nota": (
                "Cierre real local/VPS ejecutado. Reemplazó data/facturas.xlsx por un Excel limpio "
                "y actualizó cierre_trimestral_state.json. Este script aún no reemplaza SharePoint."
            ),
        },
    )
    escribir_resumen(resumen_path, manifest)

    print("-" * 100)
    print("✅ CIERRE TRIMESTRAL REAL LOCAL/VPS EJECUTADO.")
    print(f"Excel cierre contabilidad: {archivo_cerrado}")
    print(f"Respaldo activo antes de reemplazo: {respaldo_activo}")
    print(f"Nuevo data/facturas.xlsx limpio: {FACTURAS_PATH}")
    print(f"Manifest REAL: {manifest_path}")
    print(f"Resumen REAL: {resumen_path}")
    print("-" * 100)
    print("✅ data/facturas.xlsx fue reemplazado por estructura limpia.")
    print("✅ cierre_trimestral_state.json fue actualizado al siguiente trimestre.")
    print("⚠️ SharePoint todavía no se reemplaza en esta versión.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Solo diagnostica, no modifica archivos.")
    parser.add_argument("--local", action="store_true", help="Genera cierre local, sin reemplazar original.")
    parser.add_argument("--real", action="store_true", help="Ejecuta cierre real local/VPS.")
    parser.add_argument("--confirmar", default=None, help="Confirmación obligatoria para --real.")
    args = parser.parse_args()

    modos = [bool(args.dry_run), bool(args.local), bool(args.real)]
    if sum(modos) > 1:
        print("❌ Usa solo un modo: --dry-run, --local o --real.")
        return 1

    modo = "REAL" if args.real else "LOCAL" if args.local else "DRY RUN"

    print("=" * 100)
    print(f"CIERRE TRIMESTRAL FACTURAS - {modo}")
    print("=" * 100)
    print(f"Versión: {VERSION}")
    print(f"Root: {ROOT}")
    print("-" * 100)

    try:
        estado = cargar_estado()

        if args.real:
            validar_ejecucion_real(estado, args.confirmar)

        info_excel = diagnosticar_excel()
        backup = buscar_ultimo_backup_mensual()
        periodo = imprimir_plan(estado, info_excel, backup, modo)

        if args.real:
            ejecutar_real(periodo, info_excel, backup, estado)
        elif args.local:
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
