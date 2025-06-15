import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os
from zipfile import ZipFile
from io import BytesIO

st.set_page_config(page_title="Generador de CTS", layout="centered")
st.title("📄 Generador de Certificados CTS")

# 📁 Subida de archivos
plantilla_file = st.file_uploader("📄 Cargar plantilla Word (.docx)", type="docx")
excel_file = st.file_uploader("📊 Cargar archivo Excel (.xlsx)", type="xlsx")

if plantilla_file and excel_file:
    st.success("✅ Archivos cargados correctamente.")

    if st.button("🛠️ Generar certificados"):
        df = pd.read_excel(excel_file, sheet_name="Datos Empleados")
        buffer_zip = BytesIO()

        with ZipFile(buffer_zip, "w") as zipf:
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

                doc = DocxTemplate(plantilla_file)
                doc.render(context)

                nombre_archivo = f"CTS_{row['Número de documento']}_05_2025.docx"
                temp_output = BytesIO()
                doc.save(temp_output)
                temp_output.seek(0)
                zipf.writestr(nombre_archivo, temp_output.read())

        buffer_zip.seek(0)
        st.success("🎉 Certificados generados correctamente.")
        st.download_button("⬇️ Descargar ZIP", buffer_zip, file_name="certificados_cts.zip", mime="application/zip")
