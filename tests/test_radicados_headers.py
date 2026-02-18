# tests/test_radicados_headers.py
from config import RADICADOS_LOCAL_PATH, RADICADOS_SHEET_NAME, RAD_COL_ASUNTO, RAD_COL_RADICADO, RAD_COL_PROY
from services.radicados_service import cargar_mapa_radicados, buscar_radicado_y_proyecto

def main():
    print("LOCAL:", RADICADOS_LOCAL_PATH)
    print("SHEET:", RADICADOS_SHEET_NAME)
    print("HEADERS ESPERADOS:", RAD_COL_ASUNTO, "|", RAD_COL_RADICADO, "|", RAD_COL_PROY)

    mapa = cargar_mapa_radicados(force_reload=True)
    print("✅ Total filas mapeadas:", len(mapa))

    for k in ["N0761734", "0761734", "N-0761734", "Factura-N0761734"]:
        print(f"• {k} ->", mapa.get(k))


if __name__ == "__main__":
    main()
