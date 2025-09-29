# tests/sp_check.py (al inicio)
import os, sys
ROOT = os.path.dirname(os.path.abspath(__file__))  # ...\facturas_procesador\tests
ROOT = os.path.dirname(ROOT)                        # ...\facturas_procesador
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import os
from dotenv import load_dotenv
load_dotenv()

from services.m365.sp_graph import ensure_folder, upload_small_file, SP_FOLDER

print("SP_FOLDER:", SP_FOLDER)

# Ramas que usará el flujo
ramas = [
    f"{SP_FOLDER}/adjuntos",
    f"{SP_FOLDER}/extraidos",
    f"{SP_FOLDER}/excel",
]

for r in ramas:
    print("Creando/verificando:", r)
    ensure_folder(r)

# Sube un archivo de prueba a /excel
test_path = "_probar_subida.txt"
with open(test_path, "w", encoding="utf-8") as f:
    f.write("hola sharepoint")

upload_small_file(test_path, f"{SP_FOLDER}/excel/_probar_subida.txt")
print("OK: subido _probar_subida.txt a /excel")
