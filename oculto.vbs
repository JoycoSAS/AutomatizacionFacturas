Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "C:\Users\Infraestructura\Downloads\facturas_procesador\ejecutar_facturas.bat" & Chr(34), 0
Set WshShell = Nothing
