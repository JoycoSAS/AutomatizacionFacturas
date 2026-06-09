"""
MAIN INTEGRADO CON CONFIGURACIÓN ÚNICA - APROBADAS + NOTAS CRÉDITO INBOX
Versión: 2026-05-29-SOLO-EXCEL-SP-ENV-UNICO

Objetivo:
- Ejecutar en un solo ciclo el flujo normal de aprobadas y el flujo temporal de notas crédito.
- Usar el .env como ÚNICO lugar editable de configuración.
- Preparar modo producción con ventana corta.
- Permitir modo histórico/reproceso para pruebas comparativas.
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
"""

import os
import traceback
from dataclasses import dataclass

from dotenv import load_dotenv

load_dotenv()

VERSION_MAIN = "2026-05-29-CONFIG-UNICA-SOLO-EXCEL-SP-ENV-UNICO"


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


def cargar_configuracion() -> ConfigEjecucion:
    """
    Carga configuración única ya normalizada desde .env.
    No edites este main para cambiar la corrida; edita .env.
    """
    modo = _env_str("FACTURAS_MODO", EARLY_MODO).upper()
    if modo not in {"HISTORICO", "DIARIO", "PRODUCCION", "PRODUCCIÓN"}:
        print(f"⚠️ FACTURAS_MODO inválido={modo!r}. Uso PRODUCCION.")
        modo = "PRODUCCION"

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


# ============================================================
# MAIN INTEGRADO
# ============================================================

def main() -> None:
    print(f"🔥 MAIN INTEGRADO ACTIVO: {VERSION_MAIN}")

    cfg = cargar_configuracion()
    imprimir_configuracion(cfg)

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

        try:
            run_desde_aprobadas(
                max_aprobados=cfg.max_mensajes,
                max_zip_buscar=cfg.max_zip_buscar,
                since_days=cfg.since_days,
                unread_only=cfg.unread_only,
                usar_processed_store=cfg.usar_processed_store,
            )
        except Exception as e:
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
            print(f"❌ Error ejecutando flujo de notas crédito Inbox: {e}")
            traceback.print_exc()
    else:
        print("ℹ️ FACTURAS_RUN_NOTAS_CREDITO=0. Se omite flujo de notas crédito Inbox.")

    print("\n✅ MAIN INTEGRADO FINALIZADO")


if __name__ == "__main__":
    main()
