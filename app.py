import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import os
import tempfile
from io import BytesIO
from datetime import datetime
from pathlib import Path

# Título de la aplicación
st.title("Generador de Certificados CTS")

# Subida de archivos
st.markdown("### 1. Sube los archivos necesarios")
uploaded_excel = st.file_uploader("📄 Archivo Excel con datos de empleados", type=["xlsx"])
uploaded_template = st.file_uploader("📄 Plantilla Word (.docx) del certificado", type=["docx"])

# Botón para generar certificados
if uploaded_excel and uploaded_template:
    if st.button("🚀 Generar certificados"):
        with tempfile.TemporaryDirectory() as tmpdir:
            # Leer datos del Excel
            df = pd.read_excel(uploaded_excel, engine="openpyxl")

            # Cargar plantilla Word
            template_path = os.path.join(tmpdir, "plantilla.docx")
            with open(template_path, "wb") as f:
                f.write(uploaded_template.read())

            # Crear carpeta para certificados
            certificados_dir = os.path.join(tmpdir, "certificados")
            os.makedirs(certificados_dir, exist_ok=True)

            # Generar certificados
for _, row in df.iterrows():
    context = {
        'nombre': row['Nombre'],
        'dni': row['Tipo de documento'],
        'dninumero': row['Número de documento'],
        'fechaingreso': row['Fecha Ingreso'],
        'cts': row['Cuenta CTS'],
        'banco': row['Entidad CTS'],
        'base': f"S/ {row['Sueldo Base']:.2f}",
        'asfam': f"S/ {row['Asignacion Familiar']:.2f}",
        'gra': f"S/ {row['Sexto Gratificacion']:.2f}",
        'total': f"S/ {row['Base Calculo']:.2f}",
        'mes': row['Meses'],
        'mestot': f"S/ {row['Importe Meses']:.2f}",
        'dias': row['Dias'],
        'diatot': f"S/ {row['Importe Dias']:.2f}",
  'totaldep': f"S/ {row['Total CTS']:.2f}",
'importe': row['Letra'],
    }
    doc.render(context)
    doc.save(f"CTS_0{row['Número de documento']}_05_2025.docx")

import os

for archivo in os.listdir():
    if archivo.endswith(".docx") and archivo.startswith("CTS_"):
        os.system(f'libreoffice --headless --convert-to pdf "{archivo}" --outdir "."')

from zipfile import ZipFile

nombre_zip = "BOLETA.zip"

with ZipFile(nombre_zip, "w") as zipf:
    for archivo in os.listdir():
        if archivo.startswith("CTS_") and archivo.endswith(".docx"):
            zipf.write(archivo)

print("✅ ZIP creado:", nombre_zip)

from IPython.display import FileLink
FileLink(nombre_zip)

