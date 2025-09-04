import streamlit as st
import pandas as pd
from docx import Document
from openpyxl import load_workbook
import os, shutil, zipfile, tempfile

st.title("📂 Generador de Documentos - EssilorLuxottica")

# ==============================
# 1. Subir archivos base
# ==============================
clientes_file = st.file_uploader("📊 Sube PLANTILLA_DATOS.xlsx (clientes)", type=["xlsx"])
word_file = st.file_uploader("📝 Sube FR-EI-02.docx (plantilla Word)", type=["docx"])
excel_file = st.file_uploader("📑 Sube FR-EI-04.xlsx (plantilla Excel)", type=["xlsx"])

if clientes_file and word_file and excel_file:
    if st.button("🚀 Generar Documentos"):
        with st.spinner("Procesando clientes..."):
            # Crear carpeta temporal
            temp_dir = tempfile.mkdtemp()
            output_dir = os.path.join(temp_dir, "Resultados_Finales")
            os.makedirs(output_dir, exist_ok=True)

            # Leer datos de clientes
            df = pd.read_excel(clientes_file)

            # Iterar clientes
            for index, row in df.iterrows():
                cliente = str(row["CLIENTE"]).strip().replace(" ", "_")
                equipo  = str(row["NOMBRE DEL EQUIPO"]).strip().replace(" ", "_")
                referencia = str(row["REFERENCIA"]).strip() if pd.notna(row["REFERENCIA"]) else ""
                serie      = str(row["SERIE"]).strip() if pd.notna(row["SERIE"]) else ""
                fecha_word = str(row["FECHA INSTALACION(WORD)"]).strip() if pd.notna(row["FECHA INSTALACION(WORD)"]) else ""

                # Fechas
                dia  = str(row["DD"]).strip() if pd.notna(row["DD"]) else ""
                mes  = str(row["MM"]).strip() if pd.notna(row["MM"]) else ""
                anio = str(row["AA"]).strip() if pd.notna(row["AA"]) else ""

                entidad  = str(row["ENTIDAD"]).strip() if pd.notna(row["ENTIDAD"]) else ""
                ciudad   = str(row["CIUDAD"]).strip() if pd.notna(row["CIUDAD"]) else ""
                telefono = str(row["TELEFONO CLIENTE"]).strip() if pd.notna(row["TELEFONO CLIENTE"]) else ""

                # Carpeta temporal por cliente/equipo
                folder_name = f"{cliente}_{equipo}"
                folder_path = os.path.join(output_dir, folder_name)
                os.makedirs(folder_path, exist_ok=True)

                # ==============================
                # A) Generar Word
                # ==============================
                doc = Document(word_file)

                for p in doc.paragraphs:
                    if "(Nombre cliente)" in p.text:
                        p.text = p.text.replace("(Nombre cliente)", row["CLIENTE"])

                table = doc.tables[0]
                table.cell(0,1).text = str(row["NOMBRE DEL EQUIPO"])
                table.cell(1,1).text = referencia
                table.cell(2,1).text = serie
                table.cell(3,1).text = fecha_word

                word_path = os.path.join(folder_path, f"Entrega_{cliente}_{equipo}.docx")
                doc.save(word_path)

                # ==============================
                # B) Generar Excel
                # ==============================
                wb = load_workbook(excel_file)
                ws = wb.active

                ws["D9"]  = row["NOMBRE DEL EQUIPO"]
                ws["V24"] = dia
                ws["X24"] = mes
                ws["Y24"] = anio
                ws["D22"]  = entidad
                ws["AE22"] = ciudad
                ws["AE24"] = telefono
                ws["AD7"] = serie

                excel_path = os.path.join(folder_path, f"HojaVida_{cliente}_{equipo}.xlsx")
                wb.save(excel_path)

                # ==============================
                # C) Crear ZIP por cliente_equipo
                # ==============================
                zip_path = os.path.join(output_dir, f"{cliente}_{equipo}.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for f in os.listdir(folder_path):
                        zipf.write(os.path.join(folder_path, f), arcname=f)

                shutil.rmtree(folder_path)

            # ==============================
            # D) ZIP final con todo
            # ==============================
            final_zip = os.path.join(temp_dir, "Resultados_Todos.zip")
            shutil.make_archive(final_zip.replace(".zip", ""), "zip", output_dir)

            # Descargar
            with open(final_zip, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar Resultados_Todos.zip",
                    data=f,
                    file_name="Resultados_Todos.zip",
                    mime="application/zip"
                )

        st.success("✅ Proceso terminado correctamente.")
