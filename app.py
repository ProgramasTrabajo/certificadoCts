import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os

st.title("🧾 Generador de Certificados CTS")

# Cargar archivos
plantilla = "CERTIFICADO CTS.docx"
datos_excel = "Boletas_de_Pago.xlsx"

if st.button("Generar certificados"):
    df = pd.read_excel(datos_excel, sheet_name="Datos Empleados")
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
        doc = DocxTemplate(plantilla)
        nombre_archivo = f"CTS_{row['Número de documento']}_05_2025.docx"
        doc.render(context)
        doc.save(nombre_archivo)
    st.success("✅ Certificados generados. Puedes descargarlos como ZIP usando herramientas externas.")
