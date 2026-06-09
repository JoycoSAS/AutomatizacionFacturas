from controllers.aprobadas_controller import run_notas_credito_inbox_prueba

if __name__ == "__main__":
    run_notas_credito_inbox_prueba(
        max_correos=5,
        since_days=120,
        marcar_leido=False,
        usar_processed_store=False,
        aplicar_signo_nota_credito=False,
    )
