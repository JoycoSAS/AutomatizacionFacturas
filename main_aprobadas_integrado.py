"""
MAIN INTEGRADO CON CONFIGURACIÓN ÚNICA - APROBADAS + NOTAS CRÉDITO INBOX
Versión: 2026-06-09-HISTORICO-ASC-CONFIG-UNICA-SOLO-EXCEL-SP

Objetivo:
- Ejecutar en un solo ciclo el flujo normal de aprobadas y, si se activa, el flujo temporal de notas crédito.
- Usar el .env como ÚNICO lugar editable de configuración.
- Mantener producción estable con ventana corta.
- Preparar modo histórico/reproceso con orden ASC: correos antiguos primero.
- Permitir prueba histórica segura tipo DRY-RUN cuando el controller lo soporte.
- Evitar subir PDF/XML/ZIP/adjuntos/extraídos a SharePoint.
- Mantener actualización del Excel Web por Workbook API.

Regla productiva actual:
- Descargar localmente: SÍ
- Extraer localmente: SÍ
- Procesar localmente: SÍ
- Limpiar temporales locales: SÍ
- Subir PDF/XML/ZIP/adjuntos/extraídos a SharePoint: NO
- Actualizar Excel Web por Workbook API: SÍ

IMPORTANTE:
- No cambies este main para cambiar días, límites o modo.
- Cambia SOLO el archivo .env.
- Para histórico real, primero probar con FACTURAS_HISTORICO_DRY_RUN=1.
"""

import inspect
import os
import traceback
from dataclasses import dataclass
from typing import Any, Dict

from dotenv import load_dotenv

load_dotenv()

VERSION_MAIN = "2026-07-27-HISTORICO-ASC-CONFIG-UNICA-SOLO-EXCEL-SP-RC-ERRORES"


# ============================================================
# HELPERS TEMPRANOS
# ============================================================

def _raw_str(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value is None:
        return default
    value = str(value).strip()
    return value if value else default


def _raw_bool(value, default: bool = False) -> bool:
    if value is None or str(value).strip() == "":
        return bool(default)

    v = str(value).strip().lower()

    if v in {"1", "true", "yes", "y", "si", "sí", "on"}:
        return True

    if v in {"0", "false", "no", "n", "off"}:
        return False

    return bool(default)


def _raw_bool_env(name: str, default: bool = False) -> bool:
    return _raw_bool(os.getenv(name), default=default)


def _raw_int_env(name: str, default: int) -> int:
    raw = os.getenv(name)

    if raw is None or str(raw).strip() == "":
        return int(default)

    try:
        return int(str(raw).strip())
    except Exception:
        print(f"⚠️ Variable {name} inválida={raw!r}. Uso default={default}.")
        return int(default)


def _normalizar_orden_historico(raw: str) -> str:
    value = str(raw or "").strip().upper()

    if value in {"", "NONE", "NO", "0", "OFF", "FALSE"}:
        return ""

    if value in {"ASC", "ASCENDENTE", "ANTIGUOS_PRIMERO", "OLD_FIRST", "OLDEST_FIRST"}:
        return "ASC"

    if value in {"DESC", "DESCENDENTE", "RECIENTES_PRIMERO", "NEW_FIRST", "NEWEST_FIRST"}:
        return "DESC"

    print(f"⚠️ FACTURAS_ORDEN_HISTORICO inválido={raw!r}. Se deja sin orden especial.")
    return ""


# ============================================================
# CONFIGURACIÓN TEMPRANA DESDE .env
# ============================================================
# Este bloque existe porque el controller lee algunas variables al importar.
# Por eso normalizamos todo ANTES de importar controllers.aprobadas_controller.
#
# Regla:
#   .env = único lugar editable
#   main = lee .env y fuerza/normaliza defensas
#   controller = ejecuta la lógica
# ============================================================

EARLY_MODO = _raw_str("FACTURAS_MODO", "PRODUCCION").upper()

if EARLY_MODO not in {"HISTORICO", "DIARIO", "PRODUCCION", "PRODUCCIÓN"}:
    print(f"⚠️ FACTURAS_MODO inválido={EARLY_MODO!r}. Uso PRODUCCION.")
    EARLY_MODO = "PRODUCCION"

EARLY_ES_HISTORICO = EARLY_MODO == "HISTORICO"
EARLY_ES_DIARIO = EARLY_MODO in {"DIARIO", "PRODUCCION", "PRODUCCIÓN"}

if EARLY_ES_DIARIO:
    EARLY_DEFAULT_SINCE_DAYS = 6
    EARLY_DEFAULT_MAX_MENSAJES = 1000
    EARLY_DEFAULT_MAX_ZIP_BUSCAR = 1000
    EARLY_DEFAULT_UNREAD_ONLY = True
    EARLY_DEFAULT_PROCESSED_STORE = True
    EARLY_DEFAULT_MARCAR_LEIDO = False
else:
    EARLY_DEFAULT_SINCE_DAYS = 120
    EARLY_DEFAULT_MAX_MENSAJES = 5000
    EARLY_DEFAULT_MAX_ZIP_BUSCAR = 3000
    EARLY_DEFAULT_UNREAD_ONLY = False
    EARLY_DEFAULT_PROCESSED_STORE = False
    EARLY_DEFAULT_MARCAR_LEIDO = False

EARLY_SINCE_DAYS = _raw_int_env("FACTURAS_SINCE_DAYS", EARLY_DEFAULT_SINCE_DAYS)
EARLY_MAX_MENSAJES = _raw_int_env("FACTURAS_MAX_MENSAJES", EARLY_DEFAULT_MAX_MENSAJES)
EARLY_MAX_ZIP_BUSCAR = _raw_int_env("FACTURAS_MAX_ZIP_BUSCAR", EARLY_DEFAULT_MAX_ZIP_BUSCAR)
EARLY_UNREAD_ONLY = _raw_bool_env("FACTURAS_UNREAD_ONLY", EARLY_DEFAULT_UNREAD_ONLY)
EARLY_USE_PROCESSED_STORE = _raw_bool_env("FACTURAS_USE_PROCESSED_STORE", EARLY_DEFAULT_PROCESSED_STORE)
EARLY_MARCAR_LEIDO = _raw_bool_env("FACTURAS_MARCAR_LEIDO", EARLY_DEFAULT_MARCAR_LEIDO)
EARLY_FORZAR_SOLO_EXCEL_SP = _raw_bool_env("FACTURAS_FORZAR_SOLO_EXCEL_SP", True)

# ============================================================
# CONFIGURACIÓN HISTÓRICA TEMPRANA
# ============================================================
# Estas variables se normalizan antes de importar el controller para que
# cualquier módulo que las lea en import tenga los valores correctos.
# ============================================================

EARLY_ORDEN_HISTORICO = _normalizar_orden_historico(
    _raw_str("FACTURAS_ORDEN_HISTORICO", "ASC" if EARLY_ES_HISTORICO else "")
)

EARLY_PROCESAR_ANTIGUOS_PRIMERO = _raw_bool_env(
    "FACTURAS_PROCESAR_ANTIGUOS_PRIMERO",
    EARLY_ES_HISTORICO and EARLY_ORDEN_HISTORICO == "ASC",
)

if EARLY_PROCESAR_ANTIGUOS_PRIMERO and not EARLY_ORDEN_HISTORICO:
    EARLY_ORDEN_HISTORICO = "ASC"

EARLY_HISTORICO_DESDE = _raw_str("FACTURAS_HISTORICO_DESDE", "")
EARLY_HISTORICO_HASTA = _raw_str("FACTURAS_HISTORICO_HASTA", "")

# En histórico conviene desactivar autostop por defecto, porque si no puede cortar
# antes de llegar a mensajes antiguos útiles.
EARLY_DISABLE_AUTOSTOP = _raw_bool_env(
    "FACTURAS_DISABLE_AUTOSTOP",
    EARLY_ES_HISTORICO,
)

# DRY-RUN histórico: sirve para validar orden sin modificar nada.
# El controller debe soportarlo. Si aún no lo soporta, este main evita procesar.
EARLY_HISTORICO_DRY_RUN = _raw_bool_env("FACTURAS_HISTORICO_DRY_RUN", False)


# ============================================================
# DEFENSAS ANTES DE IMPORTAR EL CONTROLLER
# ============================================================

# 1) La muestra por proveedor queda desactivada siempre en este main integrado.
os.environ["MODO_MUESTRA_POR_PROVEEDOR"] = "0"
os.environ["MAX_FACTURAS_POR_PROVEEDOR"] = "999999"

# 2) Compatibilidad con variables antiguas.
#    Si alguna parte vieja del controller todavía lee MAIL_* o MAX_MESSAGES,
#    quedan sincronizadas con las variables nuevas FACTURAS_*.
os.environ["MAIL_LOOKBACK_DAYS"] = str(EARLY_SINCE_DAYS)
os.environ["MAIL_UNREAD_ONLY"] = "true" if EARLY_UNREAD_ONLY else "false"
os.environ["MAX_MESSAGES"] = str(EARLY_MAX_MENSAJES)

# 3) Modo solo Excel SharePoint.
#    Si FACTURAS_FORZAR_SOLO_EXCEL_SP=1, no se suben PDF/XML/ZIP/adjuntos/extraídos.
if EARLY_FORZAR_SOLO_EXCEL_SP:
    os.environ["SP_UPLOAD_DOCUMENTOS"] = "0"
    os.environ["SP_UPLOAD_HISTORIAL"] = "0"
    os.environ["SP_ENSURE_DOCUMENT_FOLDERS"] = "0"

# 4) Asegurar que las variables nuevas existan normalizadas en runtime.
os.environ["FACTURAS_MODO"] = EARLY_MODO
os.environ["FACTURAS_SINCE_DAYS"] = str(EARLY_SINCE_DAYS)
os.environ["FACTURAS_MAX_MENSAJES"] = str(EARLY_MAX_MENSAJES)
os.environ["FACTURAS_MAX_ZIP_BUSCAR"] = str(EARLY_MAX_ZIP_BUSCAR)
os.environ["FACTURAS_UNREAD_ONLY"] = "1" if EARLY_UNREAD_ONLY else "0"
os.environ["FACTURAS_USE_PROCESSED_STORE"] = "1" if EARLY_USE_PROCESSED_STORE else "0"
os.environ["FACTURAS_MARCAR_LEIDO"] = "1" if EARLY_MARCAR_LEIDO else "0"
os.environ["FACTURAS_FORZAR_SOLO_EXCEL_SP"] = "1" if EARLY_FORZAR_SOLO_EXCEL_SP else "0"

# 5) Variables nuevas de histórico.
os.environ["FACTURAS_ORDEN_HISTORICO"] = EARLY_ORDEN_HISTORICO
os.environ["FACTURAS_PROCESAR_ANTIGUOS_PRIMERO"] = "1" if EARLY_PROCESAR_ANTIGUOS_PRIMERO else "0"
os.environ["FACTURAS_HISTORICO_DESDE"] = EARLY_HISTORICO_DESDE
os.environ["FACTURAS_HISTORICO_HASTA"] = EARLY_HISTORICO_HASTA
os.environ["FACTURAS_DISABLE_AUTOSTOP"] = "1" if EARLY_DISABLE_AUTOSTOP else "0"
os.environ["FACTURAS_HISTORICO_DRY_RUN"] = "1" if EARLY_HISTORICO_DRY_RUN else "0"

# 6) Si el modo histórico desactiva autostop, se suben umbrales para evitar cortes tempranos.
#    Esto protege cargas históricas largas; producción normal no se toca.
if EARLY_DISABLE_AUTOSTOP:
    os.environ["AUTO_STOP_MIN_PROCESADOS"] = "999999999"
    os.environ["AUTO_STOP_SIN_NUEVOS_CONSEC"] = "999999999"
    os.environ["AUTO_STOP_SIN_MATCH_CONSEC"] = "999999999"
    os.environ["STOP_AFTER_ALREADY_PROCESSED"] = "999999999"


from controllers.aprobadas_controller import (  # noqa: E402
    run_desde_aprobadas,
    run_notas_credito_inbox_prueba,
)


# ============================================================
# HELPERS DE CONFIGURACIÓN
# ============================================================

def _env_str(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value is None:
        return default
    value = str(value).strip()
    return value if value else default


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)

    if raw is None or str(raw).strip() == "":
        return int(default)

    try:
        return int(str(raw).strip())
    except Exception:
        print(f"⚠️ Variable {name} inválida={raw!r}. Uso default={default}.")
        return int(default)


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)

    if raw is None or str(raw).strip() == "":
        return bool(default)

    value = str(raw).strip().lower()

    if value in {"1", "true", "yes", "y", "si", "sí", "on"}:
        return True

    if value in {"0", "false", "no", "n", "off"}:
        return False

    print(f"⚠️ Variable {name} inválida={raw!r}. Uso default={default}.")
    return bool(default)


def _orden_historico_env(default: str = "") -> str:
    return _normalizar_orden_historico(_env_str("FACTURAS_ORDEN_HISTORICO", default))


@dataclass(frozen=True)
class ConfigEjecucion:
    modo: str
    run_aprobadas: bool
    run_notas_credito: bool
    since_days: int
    max_mensajes: int
    max_zip_buscar: int
    unread_only: bool
    usar_processed_store: bool
    marcar_leido: bool
    nota_credito_valores_negativos: bool
    forzar_solo_excel_sp: bool
    sp_upload_documentos: bool
    sp_upload_historial: bool
    sp_ensure_document_folders: bool

    # Histórico / reproceso
    orden_historico: str
    procesar_antiguos_primero: bool
    historico_desde: str
    historico_hasta: str
    disable_autostop: bool
    historico_dry_run: bool


def cargar_configuracion() -> ConfigEjecucion:
    """
    Carga configuración única ya normalizada desde .env.
    No edites este main para cambiar la corrida; edita .env.
    """
    modo = _env_str("FACTURAS_MODO", EARLY_MODO).upper()

    if modo not in {"HISTORICO", "DIARIO", "PRODUCCION", "PRODUCCIÓN"}:
        print(f"⚠️ FACTURAS_MODO inválido={modo!r}. Uso PRODUCCION.")
        modo = "PRODUCCION"

    es_historico = modo == "HISTORICO"
    es_diario = modo in {"DIARIO", "PRODUCCION", "PRODUCCIÓN"}

    if es_diario:
        default_since_days = 6
        default_max_mensajes = 1000
        default_max_zip_buscar = 1000
        default_unread_only = True
        default_processed_store = True
        default_marcar_leido = False
    else:
        default_since_days = 120
        default_max_mensajes = 5000
        default_max_zip_buscar = 3000
        default_unread_only = False
        default_processed_store = False
        default_marcar_leido = False

    orden_historico = _orden_historico_env("ASC" if es_historico else "")
    procesar_antiguos_primero = _env_bool(
        "FACTURAS_PROCESAR_ANTIGUOS_PRIMERO",
        es_historico and orden_historico == "ASC",
    )

    if procesar_antiguos_primero and not orden_historico:
        orden_historico = "ASC"

    return ConfigEjecucion(
        modo=modo,
        run_aprobadas=_env_bool("FACTURAS_RUN_APROBADAS", True),
        run_notas_credito=_env_bool("FACTURAS_RUN_NOTAS_CREDITO", True),
        since_days=_env_int("FACTURAS_SINCE_DAYS", default_since_days),
        max_mensajes=_env_int("FACTURAS_MAX_MENSAJES", default_max_mensajes),
        max_zip_buscar=_env_int("FACTURAS_MAX_ZIP_BUSCAR", default_max_zip_buscar),
        unread_only=_env_bool("FACTURAS_UNREAD_ONLY", default_unread_only),
        usar_processed_store=_env_bool("FACTURAS_USE_PROCESSED_STORE", default_processed_store),
        marcar_leido=_env_bool("FACTURAS_MARCAR_LEIDO", default_marcar_leido),
        nota_credito_valores_negativos=_env_bool("FACTURAS_NOTA_CREDITO_VALORES_NEGATIVOS", False),
        forzar_solo_excel_sp=_env_bool("FACTURAS_FORZAR_SOLO_EXCEL_SP", True),
        sp_upload_documentos=_env_bool("SP_UPLOAD_DOCUMENTOS", False),
        sp_upload_historial=_env_bool("SP_UPLOAD_HISTORIAL", False),
        sp_ensure_document_folders=_env_bool("SP_ENSURE_DOCUMENT_FOLDERS", False),

        orden_historico=orden_historico,
        procesar_antiguos_primero=procesar_antiguos_primero,
        historico_desde=_env_str("FACTURAS_HISTORICO_DESDE", ""),
        historico_hasta=_env_str("FACTURAS_HISTORICO_HASTA", ""),
        disable_autostop=_env_bool("FACTURAS_DISABLE_AUTOSTOP", es_historico),
        historico_dry_run=_env_bool("FACTURAS_HISTORICO_DRY_RUN", False),
    )


def imprimir_configuracion(cfg: ConfigEjecucion) -> None:
    print("\n====================================================")
    print("⚙️ CONFIGURACIÓN GLOBAL ÚNICA DESDE .env")
    print("====================================================")
    print(f"VERSION_MAIN={VERSION_MAIN}")
    print(f"FACTURAS_MODO={cfg.modo}")
    print(f"FACTURAS_RUN_APROBADAS={cfg.run_aprobadas}")
    print(f"FACTURAS_RUN_NOTAS_CREDITO={cfg.run_notas_credito}")
    print(f"FACTURAS_SINCE_DAYS={cfg.since_days}")
    print(f"FACTURAS_MAX_MENSAJES={cfg.max_mensajes}")
    print(f"FACTURAS_MAX_ZIP_BUSCAR={cfg.max_zip_buscar}")
    print(f"FACTURAS_UNREAD_ONLY={cfg.unread_only}")
    print(f"FACTURAS_USE_PROCESSED_STORE={cfg.usar_processed_store}")
    print(f"FACTURAS_MARCAR_LEIDO={cfg.marcar_leido}")
    print(f"FACTURAS_NOTA_CREDITO_VALORES_NEGATIVOS={cfg.nota_credito_valores_negativos}")

    print("----------------------------------------------------")
    print("🧭 HISTÓRICO / REPROCESO")
    print("----------------------------------------------------")
    print(f"FACTURAS_ORDEN_HISTORICO={cfg.orden_historico or '(sin orden especial)'}")
    print(f"FACTURAS_PROCESAR_ANTIGUOS_PRIMERO={cfg.procesar_antiguos_primero}")
    print(f"FACTURAS_HISTORICO_DESDE={cfg.historico_desde or '(sin fecha desde)'}")
    print(f"FACTURAS_HISTORICO_HASTA={cfg.historico_hasta or '(sin fecha hasta)'}")
    print(f"FACTURAS_DISABLE_AUTOSTOP={cfg.disable_autostop}")
    print(f"FACTURAS_HISTORICO_DRY_RUN={cfg.historico_dry_run}")

    if cfg.modo == "HISTORICO" and cfg.orden_historico == "ASC":
        print("✅ Modo histórico ASC activo: se preparará procesamiento de antiguo → reciente.")
    elif cfg.modo == "HISTORICO":
        print("⚠️ Modo histórico activo sin ASC. Revisa FACTURAS_ORDEN_HISTORICO.")

    if cfg.historico_dry_run:
        print("🧪 DRY-RUN histórico activo: no debe modificar Excel ni marcar correos.")

    print("----------------------------------------------------")
    print("🧪 MUESTRA POR PROVEEDOR")
    print("----------------------------------------------------")
    print("MODO_MUESTRA_POR_PROVEEDOR=0")
    print("MAX_FACTURAS_POR_PROVEEDOR=999999")

    print("----------------------------------------------------")
    print("📧 COMPATIBILIDAD VARIABLES ANTIGUAS")
    print("----------------------------------------------------")
    print(f"MAIL_LOOKBACK_DAYS={os.getenv('MAIL_LOOKBACK_DAYS')}")
    print(f"MAIL_UNREAD_ONLY={os.getenv('MAIL_UNREAD_ONLY')}")
    print(f"MAX_MESSAGES={os.getenv('MAX_MESSAGES')}")

    print("----------------------------------------------------")
    print("⛔ AUTO STOP")
    print("----------------------------------------------------")
    print(f"AUTO_STOP_MIN_PROCESADOS={os.getenv('AUTO_STOP_MIN_PROCESADOS')}")
    print(f"AUTO_STOP_SIN_NUEVOS_CONSEC={os.getenv('AUTO_STOP_SIN_NUEVOS_CONSEC')}")
    print(f"AUTO_STOP_SIN_MATCH_CONSEC={os.getenv('AUTO_STOP_SIN_MATCH_CONSEC')}")
    print(f"STOP_AFTER_ALREADY_PROCESSED={os.getenv('STOP_AFTER_ALREADY_PROCESSED')}")

    if cfg.disable_autostop:
        print("✅ Autostop desactivado/neutralizado para evitar cortes tempranos.")

    print("----------------------------------------------------")
    print("📎 SHAREPOINT DOCUMENTOS")
    print("----------------------------------------------------")
    print(f"FACTURAS_FORZAR_SOLO_EXCEL_SP={cfg.forzar_solo_excel_sp}")
    print(f"SP_UPLOAD_DOCUMENTOS={cfg.sp_upload_documentos}")
    print(f"SP_UPLOAD_HISTORIAL={cfg.sp_upload_historial}")
    print(f"SP_ENSURE_DOCUMENT_FOLDERS={cfg.sp_ensure_document_folders}")

    if cfg.forzar_solo_excel_sp and not cfg.sp_upload_documentos:
        print("✅ Modo solo Excel activo: NO se subirán PDF/XML/ZIP/adjuntos/extraídos a SharePoint.")
    elif not cfg.sp_upload_documentos:
        print("✅ SP_UPLOAD_DOCUMENTOS=0: NO se subirán documentos pesados a SharePoint.")
    else:
        print("⚠️ SP_UPLOAD_DOCUMENTOS=True: se podrían subir documentos pesados a SharePoint.")

    print("====================================================\n")


def _run_desde_aprobadas_compatible(cfg: ConfigEjecucion) -> None:
    """
    Ejecuta run_desde_aprobadas manteniendo compatibilidad con controllers viejos.

    Si el controller nuevo acepta parámetros históricos, se los pasa.
    Si todavía no los acepta, no rompe producción.

    Protección especial:
    - Si FACTURAS_HISTORICO_DRY_RUN=1 pero el controller no soporta historico_dry_run,
      este main NO ejecuta procesamiento real para evitar cambios accidentales.
    """
    kwargs: Dict[str, Any] = {
        "max_aprobados": cfg.max_mensajes,
        "max_zip_buscar": cfg.max_zip_buscar,
        "since_days": cfg.since_days,
        "unread_only": cfg.unread_only,
        "usar_processed_store": cfg.usar_processed_store,
    }

    try:
        sig = inspect.signature(run_desde_aprobadas)
        params = set(sig.parameters.keys())
    except Exception:
        params = set(kwargs.keys())

    extras: Dict[str, Any] = {
        "orden_historico": cfg.orden_historico,
        "procesar_antiguos_primero": cfg.procesar_antiguos_primero,
        "historico_desde": cfg.historico_desde,
        "historico_hasta": cfg.historico_hasta,
        "disable_autostop": cfg.disable_autostop,
        "historico_dry_run": cfg.historico_dry_run,
    }

    for key, value in extras.items():
        if key in params:
            kwargs[key] = value

    if cfg.historico_dry_run and "historico_dry_run" not in params:
        print("🛑 FACTURAS_HISTORICO_DRY_RUN=1, pero el controller actual todavía no soporta historico_dry_run.")
        print("   Por seguridad NO se ejecuta procesamiento real.")
        print("   Siguiente paso: corregir controllers\\aprobadas_controller.py para soportar DRY-RUN histórico.")
        return

    run_desde_aprobadas(**kwargs)


# ============================================================
# MAIN INTEGRADO
# ============================================================

def main() -> int:
    print(f"🔥 MAIN INTEGRADO ACTIVO: {VERSION_MAIN}")

    cfg = cargar_configuracion()
    imprimir_configuracion(cfg)

    errores: list[str] = []

    # ============================================================
    # 1) FLUJO NORMAL COMPLETO DE APROBADAS
    # ============================================================
    if cfg.run_aprobadas:
        print("\n====================================================")
        print("🚀 INICIANDO FLUJO 1/2: APROBADAS NORMAL")
        print("====================================================")
        print("Usa la MISMA configuración global desde .env:")
        print(f"  max_aprobados={cfg.max_mensajes}")
        print(f"  max_zip_buscar={cfg.max_zip_buscar}")
        print(f"  since_days={cfg.since_days}")
        print(f"  unread_only={cfg.unread_only}")
        print(f"  usar_processed_store={cfg.usar_processed_store}")

        if cfg.modo == "HISTORICO":
            print("----------------------------------------------------")
            print("🧭 Parámetros históricos")
            print("----------------------------------------------------")
            print(f"  orden_historico={cfg.orden_historico or '(sin orden especial)'}")
            print(f"  procesar_antiguos_primero={cfg.procesar_antiguos_primero}")
            print(f"  historico_desde={cfg.historico_desde or '(sin fecha desde)'}")
            print(f"  historico_hasta={cfg.historico_hasta or '(sin fecha hasta)'}")
            print(f"  disable_autostop={cfg.disable_autostop}")
            print(f"  historico_dry_run={cfg.historico_dry_run}")

        try:
            _run_desde_aprobadas_compatible(cfg)
        except Exception as e:
            mensaje = (
                "Flujo de aprobadas: "
                f"{type(e).__name__}: {e}"
            )
            errores.append(mensaje)
            print(f"❌ Error ejecutando flujo de aprobadas: {e}")
            traceback.print_exc()
    else:
        print("ℹ️ FACTURAS_RUN_APROBADAS=0. Se omite flujo de aprobadas.")

    # ============================================================
    # 2) FLUJO TEMPORAL INTEGRADO: NOTAS CRÉDITO EN INBOX
    # ============================================================
    if cfg.run_notas_credito:
        print("\n====================================================")
        print("🚀 INICIANDO FLUJO 2/2: NOTAS CRÉDITO INBOX")
        print("====================================================")
        print("Usa la MISMA configuración global desde .env:")
        print(f"  max_correos={cfg.max_mensajes}")
        print(f"  since_days={cfg.since_days}")
        print(f"  marcar_leido={cfg.marcar_leido}")
        print(f"  usar_processed_store={cfg.usar_processed_store}")
        print(f"  aplicar_signo_nota_credito={cfg.nota_credito_valores_negativos}")

        try:
            run_notas_credito_inbox_prueba(
                max_correos=cfg.max_mensajes,
                since_days=cfg.since_days,
                marcar_leido=cfg.marcar_leido,
                usar_processed_store=cfg.usar_processed_store,
                aplicar_signo_nota_credito=cfg.nota_credito_valores_negativos,
            )
        except Exception as e:
            mensaje = (
                "Flujo de notas crédito Inbox: "
                f"{type(e).__name__}: {e}"
            )
            errores.append(mensaje)
            print(f"❌ Error ejecutando flujo de notas crédito Inbox: {e}")
            traceback.print_exc()
    else:
        print("ℹ️ FACTURAS_RUN_NOTAS_CREDITO=0. Se omite flujo de notas crédito Inbox.")

    if errores:
        print("\n❌ MAIN INTEGRADO FINALIZADO CON ERRORES")
        for error in errores:
            print(f"   - {error}")
        return 1

    print("\n✅ MAIN INTEGRADO FINALIZADO")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())