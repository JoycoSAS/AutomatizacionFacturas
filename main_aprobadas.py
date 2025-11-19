# main_aprobadas.py
from controllers.aprobadas_controller import run_desde_aprobadas

if __name__ == "__main__":
    print(">> Iniciando flujo por carpeta 'Facturas aprobadas'...")
    # Puedes ajustar estos parámetros si lo necesitas más adelante
    run_desde_aprobadas(max_aprobados=50, max_zip_buscar=100)
