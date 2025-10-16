import os
from pathlib import Path

def borrar_pdfs_en_arbol(carpeta: str) -> int:
    """
    Elimina TODOS los .pdf/.PDF bajo 'carpeta' (incluye subcarpetas).
    Devuelve la cantidad realmente borrada. No toca XML/ZIP ni borra carpetas.
    Deja trazas de depuración para verificar ruta y candidatos.
    """
    base = Path(carpeta)
    if not base.exists():
        print(f"[DEBUG] Limpieza: carpeta NO existe -> {carpeta}")
        return 0

    # Buscar candidatos en minúscula y mayúscula (por si el FS es case sensitive)
    cand  = list(base.rglob("*.pdf"))
    cand += list(base.rglob("*.PDF"))

    print(f"[DEBUG] Limpieza en: {base.resolve()} | candidatos: {len(cand)}")
    borrados = 0

    for p in cand:
        try:
            # Quitar posible atributo de solo lectura (Windows)
            try:
                os.chmod(p, 0o666)
            except Exception:
                pass
            p.unlink(missing_ok=True)
            borrados += 1
        except Exception as e:
            # No detenemos el flujo si alguno no se puede borrar (bloqueado/en uso)
            print(f"[DEBUG] No se pudo borrar '{p}': {e}")

    return borrados
