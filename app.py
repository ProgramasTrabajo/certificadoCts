import pandas as pd
from docxtpl import DocxTemplate
import os
import subprocess
from zipfile import ZipFile
from datetime import datetime

# Rutas
archivo_excel = "Boletas_de_Pago.xlsx"
plantilla_word = "CERTIFICADO CTS.docx"
output_folder = "certificados"
pdf_folder = "certificados_pdf"
zip_name = "certificados_pdf.zip"

# Crear carpetas si no existen
os.makedirs(output_folder, exist_ok=True)
os.makedirs(pdf_folder, exist_ok=True)

# Leer Excel
df = pd.read_excel(archivo_excel, sheet_name="Datos Empleados")

# Generar archivos Word
for _, row in df.iterrows():
    context = {
        'nombre': row['Nombre'],
        'dni': row['Tipo de documento'],
        'dninumero': str(row['Número de documento']).zfill(8),
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

    doc = DocxTemplate(plantilla_word)
    filename = f"CTS_0{str(row['Número de documento']).zfill(8)}_05_2025.docx"
    ruta_docx = os.path.join(output_folder, nombre_archivo)
    doc.render(context)
    doc.save(ruta_docx)

    # Convertir a PDF con LibreOffice
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", pdf_folder,
        ruta_docx
    ])

# Crear ZIP con todos los PDFs
with ZipFile(zip_name, "w") as zipf:
    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            zipf.write(os.path.join(pdf_folder, filename), filename)

print(f"✅ Certificados generados y convertidos a PDF. ZIP listo: {zip_name}")
