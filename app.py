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

!apt-get install -y libreoffice

import os

for archivo in os.listdir():
    if archivo.endswith(".docx") and archivo.startswith("CTS_"):
        os.system(f'libreoffice --headless --convert-to pdf "{archivo}" --outdir "."')

            # Comprimir en ZIP
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in os.listdir(certificados_dir):
                    file_path = os.path.join(certificados_dir, file)
                    zipf.write(file_path, arcname=file)

            zip_buffer.seek(0)

            # Descargar ZIP
            st.success("✅ Certificados generados correctamente.")
            st.download_button(
                label="📦 Descargar ZIP con certificados",
                data=zip_buffer,
                file_name=f"certificados_cts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
else:
    st.info("Por favor, sube el archivo Excel y la plantilla Word para continuar.")
