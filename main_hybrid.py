# main_hybrid.py
# Runner mínimo para ejecutar el pipeline sin argumentos manuales.
# Mantiene tu flujo tal cual (leer correo, ZIP/XML, Excel, mover a 'procesados').

from controllers.cloud_pipeline import run_hibrido

# Ajusta estos valores si quieres probar con menos/más correos
MAX_MESSAGES = 2000    # cuántos correos revisar en este disparo
SINCE_DAYS   = 200     # ventana en días hacia atrás (opcional, puede ser None)

if __name__ == "__main__":
    print(">> Iniciando pipeline híbrido...")
    run_hibrido(max_messages=MAX_MESSAGES, since_days=SINCE_DAYS)
