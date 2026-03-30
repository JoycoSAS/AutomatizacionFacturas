# main_aprobadas.py
from controllers.aprobadas_controller import run_desde_aprobadas

if __name__ == "__main__":
    print(">> Iniciando flujo por carpeta 'Facturas aprobadas' (TEST AIDX)...")
    run_desde_aprobadas(
        max_aprobados=100,
        max_zip_buscar=200,  
        
    )
