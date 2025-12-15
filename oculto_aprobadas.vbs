Set WshShell = CreateObject("WScript.Shell")
' Ruta completa al .bat
bat = "C:\Users\Infraestructura\Downloads\facturas_procesador\ejecutar_aprobadas.bat"

' Ejecutar oculto (0 = ventana oculta, True = esperar a que termine)
WshShell.Run """" & bat & """", 0, True
