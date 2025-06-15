import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from zipfile import ZipFile
from io import BytesIO

st.set_page_config(page_title="Generador CTS", layout="centered")
st.title("📄 Generador de Certificados CTS")

plantilla_file = st.file_uploader("📄 Subir plantilla Word (.docx)", type="docx")
excel_file = st.file_uploader("📊 Subir archivo Excel (.xlsx)", type="xlsx")

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
                    'dninumero': str(row['Número de documento']).zfill(9),
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

                output = BytesIO()
                doc.save(output)
                output.seek(0)
                filename = f"CTS_0{str(row['Número de documento']).zfill(8)}_05_2025.docx"
                zipf.writestr(filename, output.read())

        buffer_zip.seek(0)
        st.download_button("⬇️ Descargar certificados en ZIP", buffer_zip, file_name="certificados_cts.zip", mime="application/zip")
