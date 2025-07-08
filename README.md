# PlantillaPDF_paraListas_leadsCon_ReportLab

Este repositorio contiene un pequeño script de Python que genera plantillas PDF estilizadas a partir de un archivo CSV de leads. Utiliza ReportLab para el diseño del documento y cuenta con una interfaz gráfica sencilla para seleccionar el CSV y crear el reporte.

## Instalación

1. (Opcional) Crea y activa un entorno virtual:

```bash
python3 -m venv venv
source venv/bin/activate
```

2. Instala las dependencias necesarias:

```bash
pip install pandas reportlab
```

`tkinter` suele venir incluido con Python.

## Uso

Ejecuta el script y sigue las instrucciones que aparecen en pantalla:

```bash
python generate_styled_pdf_template_with_ui.py
```

Tras elegir tu archivo CSV, se generará un PDF con el nombre `reporte_leads_estilizado.pdf` en la misma carpeta.
