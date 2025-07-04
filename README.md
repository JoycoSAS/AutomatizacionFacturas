# Procesador de Facturas Electrónicas

Este proyecto automatiza la descarga, extracción y procesamiento de facturas electrónicas contenidas en archivos `.zip` recibidos por correo Outlook. Extrae los datos relevantes desde los archivos XML y los almacena en un archivo Excel, evitando duplicados por número de factura.

---

## Características

* Se conecta a una bandeja de entrada Outlook.
* Detecta correos con archivos `.zip` que contengan `.xml`.
* Guarda y descomprime los `.zip` en carpetas organizadas.
* Extrae los datos del XML (como emisor, NIT, cliente, valores de IVA, total, etc).
* Actualiza un archivo Excel con los datos de nuevas facturas.
* Registra un historial de ejecuciones con errores y archivos procesados.

---

## Estructura del Proyecto

```
facturas_procesador/
├── controllers/
│   └── procesador_controller.py
├── services/
│   ├── correo_service.py
│   ├── zip_service.py
│   ├── factura_service.py
│   └── excel_service.py
├── utils/
│   ├── helpers.py
│   └── errores.py
├── data/
│   ├── adjuntos/
│   ├── extraidos/
│   ├── facturas.xlsx
│   └── historial_ejecuciones.xlsx
├── config.py
├── main.py
├── requirements.txt
└── README.md
```

---

## Requisitos

* Windows con Outlook instalado y cuenta configurada.
* Python 3.10+
* Librerías del archivo `requirements.txt`

---

## Ejecución

1. Instalar dependencias:

```bash
pip install -r requirements.txt
```

2. Ejecutar el script:

```bash
python main.py
```

Esto iniciará la lectura de correos, descarga de ZIPs, extracción de XMLs y actualización de Excel.

---

## Personalización

Puedes modificar el correo de origen (`STORE_NAME`) o las rutas en `config.py` según tu organización o entorno local.

---

## Notas

* Solo se procesan archivos ZIP que contengan al menos un XML.
* Las facturas se identifican por el campo "Número de factura" para evitar duplicados.
* A futuro puede integrarse una base de datos o procesamiento masivo en la nube.

---

## Autor

Infraestructura TI - Joyco

Daniel Andres Leones Posso Ingeniero de Software
