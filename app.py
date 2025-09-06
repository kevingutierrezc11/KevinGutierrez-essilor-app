import io
import os
import shutil
import zipfile
import tempfile

import pandas as pd
import streamlit as st
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import streamlit as st

# Mostrar imagen centrada
st.markdown(
    """
    <div style="text-align: center;">
        <img src="logo.png" width="250">
    </div>
    """,
    unsafe_allow_html=True
)
st.set_page_config(page_title="Generador DOCUMENTACIÓN CLIENTES - EssilorLuxottica", layout="wide")
st.title("📂 Generador DOCUMENTACIÓN CLIENTES — versión web - Auor : KEVIN EDUARDO GUTIERREZ CASTILLO ")

st.markdown(
    """
    **Instrucciones rápidas**:
    1. Sube **PLANTILLA_DATOS.xlsx** (Excel con filas de clientes).
    2. Sube **FR-EI-02.docx** (plantilla Word).
    3. Sube **FR-EI-04.xlsx**, **FR-EI-03.xlsx**, **FR-EI-05.xlsx** (plantillas Excel: Hoja de vida, Protocolo, Cronograma).
    4. Haz clic en **Generar** — se descargará `DOCUMENTACION_CLIENTES.zip` que contiene un ZIP por cliente/equipo .
    """
)

# -------------------------
# Upload inputs
# -------------------------
st.header("1) Cargar archivos base")
col1, col2 = st.columns([2, 1])

with col1:
    plantilla_datos_file = st.file_uploader("📊 PLANTILLA_DATOS.xlsx (Excel con clientes)", type=["xlsx", "xls", "csv"])
    fr_ei_02_file = st.file_uploader("📝 FR-EI-02.docx (plantilla Word)", type=["docx"])
    fr_ei_04_file = st.file_uploader("📑 FR-EI-04.xlsx (plantilla Hoja de Vida)", type=["xlsx"])
    fr_ei_03_file = st.file_uploader("📑 FR-EI-03.xlsx (plantilla Protocolo MMTO preventivo)", type=["xlsx"])
    cronograma_file = st.file_uploader("📑 FR-EI-05.xlsx (plantilla Cronograma de mantenimiento)", type=["xlsx"])
with col2:
    st.markdown("**Estado de carga**")
    st.write("PLANTILLA_DATOS:", getattr(plantilla_datos_file, "name", "—"))
    st.write("FR-EI-02 (Word):", getattr(fr_ei_02_file, "name", "—"))
    st.write("FR-EI-04 (Excel):", getattr(fr_ei_04_file, "name", "—"))
    st.write("FR-EI-03 (Excel):", getattr(fr_ei_03_file, "name", "—"))
    st.write("FR-EI-05 (Excel):", getattr(cronograma_file, "name", "—"))

st.markdown("---")

# -------------------------
# Helpers
# -------------------------
def write_respecting_merged(ws, target_cell, value):
    """
    Si target_cell pertenece a un rango mergeado, escribe en la celda top-left del rango.
    Si no, escribe directamente.
    """
    for merged in ws.merged_cells.ranges:
        if target_cell in str(merged):
            tl_col = get_column_letter(merged.min_col)
            tl_row = merged.min_row
            top_left = f"{tl_col}{tl_row}"
            ws[top_left] = value
            return
    ws[target_cell] = value

def safe_str(val):
    return "" if pd.isna(val) else str(val)

# -------------------------
# Validación mínima
# -------------------------
if not (plantilla_datos_file and fr_ei_02_file and fr_ei_04_file and fr_ei_03_file and cronograma_file):
    st.info("Sube todos los archivos listados arriba para habilitar la generación.")
    st.stop()

# -------------------------
# Leer plantilla_datos en DataFrame
# -------------------------
try:
    if plantilla_datos_file.name.lower().endswith(".csv"):
        df = pd.read_csv(plantilla_datos_file)
    else:
        df = pd.read_excel(plantilla_datos_file, engine="openpyxl")
except Exception as e:
    st.error(f"No se pudo leer PLANTILLA_DATOS: {e}")
    st.stop()

st.subheader("Vista previa (primeras filas) — revisa que los nombres de columna coincidan con tu notebook")
st.dataframe(df.head(10))

# Leer bytes de plantillas (para poder reutilizarlas múltiples veces)
fr_ei_02_bytes = fr_ei_02_file.read()
fr_ei_04_bytes = fr_ei_04_file.read()
fr_ei_03_bytes = fr_ei_03_file.read()
cronograma_bytes = cronograma_file.read()

# -------------------------
# Botón de generación
# -------------------------
st.header("2) Generar documentación")
base_name = st.text_input("Nombre base para archivos/ZIP (opcional)", value="DOCUMENTACION_CLIENTES")
generate = st.button("🚀 Generar DOCUMENTACIÓN (igual al cuaderno)")

if generate:
    with st.spinner("Generando archivos — esto puede tardar según la cantidad de filas..."):
        # carpeta temporal
        tmp_root = tempfile.mkdtemp()
        output_dir = os.path.join(tmp_root, "DOCUMENTACION_CLIENTES")
        os.makedirs(output_dir, exist_ok=True)

        created_zip_paths = []

        # Iterar sobre filas
        for index, row in df.iterrows():
            try:
                cliente = safe_str(row.get("CLIENTE", "")).strip().replace(" ", "_")
                equipo = safe_str(row.get("NOMBRE DEL EQUIPO", "")).strip().replace(" ", "_")
                referencia = safe_str(row.get("REFERENCIA", ""))
                serie = safe_str(row.get("SERIE", ""))
                fecha_word = safe_str(row.get("FECHA INSTALACION(WORD)", ""))
                nit_cliente = safe_str(row.get("NIT CLIENTE", ""))
                tipo_mmto = safe_str(row.get("TIPO MMTO", ""))
                FRECUENCIA = safe_str(row.get("FRECUENCIA", ""))
                DIRECCION = safe_str(row.get("DIRECCION", ""))
                modelo = safe_str(row.get("MODELO", ""))
                UBICACION_DEL_EQUIPO = safe_str(row.get("UBICACIÓN DEL EQUIPO (ÁREA)", ""))
                dia = safe_str(row.get("DD", ""))
                mes = safe_str(row.get("MM", ""))
                anio = safe_str(row.get("AA", ""))
                entidad = safe_str(row.get("ENTIDAD", ""))
                ciudad = safe_str(row.get("CIUDAD", ""))
                telefono = safe_str(row.get("TELEFONO CLIENTE", ""))

                # nombre carpeta y ruta
                folder_name = f"{cliente}_{equipo}" if cliente or equipo else f"fila_{index}"
                folder_path = os.path.join(output_dir, folder_name)
                os.makedirs(folder_path, exist_ok=True)

                # ---------- A) FR-EI-02 Word ----------
                # Abrir plantilla desde bytes (copia por cada iteración)
                doc = Document(io.BytesIO(fr_ei_02_bytes))

                # Reemplazo de texto simple (igual que en tu notebook)
                for p in doc.paragraphs:
                    if "(Nombre cliente)" in p.text:
                        p.text = p.text.replace("(Nombre cliente)", safe_str(row.get("CLIENTE", "")))

                # Llenar tabla (igual índices que en notebook)
                try:
                    table = doc.tables[0]
                    table.cell(0, 1).text = safe_str(row.get("NOMBRE DEL EQUIPO", ""))
                    table.cell(1, 1).text = referencia
                    table.cell(2, 1).text = serie
                    table.cell(3, 1).text = fecha_word
                except Exception as e:
                    # Si falla por estructura de tabla, mostrar advertencia pero continuar
                    st.warning(f"Advertencia: no se pudo llenar la tabla en FR-EI-02 para fila {index}: {e}")

                word_name = f"FR-EI-02 ENTREGA DE EQUIPO A CONFORMIDAD Y CONDICIONES DE GARANTÍA-{cliente}-{equipo}.docx"
                word_path = os.path.join(folder_path, word_name)
                doc.save(word_path)

                # ---------- B) FR-EI-04 Hoja de vida (Excel) ----------
                wb_hv = load_workbook(filename=io.BytesIO(fr_ei_04_bytes))
                ws_hv = wb_hv.active

                # Escribir celdas exactas (tal como en tu notebook)
                try:
                    ws_hv["D9"] = row.get("NOMBRE DEL EQUIPO", "")
                    ws_hv["V24"] = dia
                    ws_hv["X24"] = mes
                    ws_hv["Y24"] = anio
                    ws_hv["D22"] = entidad
                    ws_hv["AE22"] = ciudad
                    ws_hv["AE24"] = telefono
                    ws_hv["AD7"] = serie
                except Exception as e:
                    st.warning(f"Advertencia: error escribiendo FR-EI-04 para fila {index}: {e}")

                excel_hv_name = f"FR-EI-04 HOJA DE VIDA DEL EQUIPO-{cliente}-{equipo}.xlsx"
                excel_hv_path = os.path.join(folder_path, excel_hv_name)
                wb_hv.save(excel_hv_path)

                # ---------- C) FR-EI-03 Protocolo MMTO ----------
                wb_mmto = load_workbook(filename=io.BytesIO(fr_ei_03_bytes))
                ws_mmto = wb_mmto.active
                try:
                    ws_mmto["A12"] = row.get("NOMBRE DEL EQUIPO", "")
                except Exception as e:
                    st.warning(f"Advertencia: error escribiendo FR-EI-03 para fila {index}: {e}")
                excel_mmto_name = f"FR-EI-03 PROTOCOLO DE MANTENIMIENTO PREVENTIVO-{cliente}-{equipo}.xlsx"
                excel_mmto_path = os.path.join(folder_path, excel_mmto_name)
                wb_mmto.save(excel_mmto_path)

                # ---------- D) FR-EI-05 Cronograma (usar función que respeta merges) ----------
                wb_crono = load_workbook(filename=io.BytesIO(cronograma_bytes))
                ws_crono = wb_crono.active
                try:
                    write_respecting_merged(ws_crono, "B10", row.get("NOMBRE DEL EQUIPO", ""))
                    write_respecting_merged(ws_crono, "E10", serie)
                    write_respecting_merged(ws_crono, "F10", tipo_mmto)
                    write_respecting_merged(ws_crono, "D5", nit_cliente)
                    write_respecting_merged(ws_crono, "R6", fecha_word)
                    write_respecting_merged(ws_crono, "G10", FRECUENCIA)
                    write_respecting_merged(ws_crono, "F5", DIRECCION)
                    write_respecting_merged(ws_crono, "R5", ciudad)
                    write_respecting_merged(ws_crono, "D10", modelo)
                    write_respecting_merged(ws_crono, "D6", UBICACION_DEL_EQUIPO)
                    write_respecting_merged(ws_crono, "D4", row.get("CLIENTE", ""))
                    write_respecting_merged(ws_crono, "R4", row.get("TELEFONO CLIENTE", ""))
                except Exception as e:
                    st.warning(f"Advertencia: error escribiendo FR-EI-05 para fila {index}: {e}")

                excel_crono_name = f"FR-EI-05 CRONOGRAMA DE MANTENIMIENTO-{cliente}-{equipo}.xlsx"
                excel_crono_path = os.path.join(folder_path, excel_crono_name)
                wb_crono.save(excel_crono_path)

                # ---------- E) ZIP por cliente_equipo ----------
                zip_path = os.path.join(output_dir, f"{cliente}_{equipo}.zip")
                with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
                    for fname in os.listdir(folder_path):
                        zipf.write(os.path.join(folder_path, fname), arcname=fname)

                # limpiar carpeta temporal del cliente
                shutil.rmtree(folder_path)

                created_zip_paths.append(zip_path)

            except Exception as e:
                st.error(f"Error procesando fila {index}: {e}")
                # intentar continuar con la siguiente fila
                continue

        # ---------- Final: ZIP con todo ----------
        final_zip_path = os.path.join(tmp_root, f"{base_name}.zip")
        with zipfile.ZipFile(final_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for p in created_zip_paths:
                zf.write(p, arcname=os.path.basename(p))

        # Leer en memoria y ofrecer descarga directa
        with open(final_zip_path, "rb") as f:
            data_bytes = f.read()

        st.success("✅ Proceso terminado. Descarga el ZIP con todos los clientes.")
        st.download_button("⬇️ Descargar DOCUMENTACION_CLIENTES.zip", data=data_bytes, file_name=f"{base_name}.zip", mime="application/zip")

        # limpieza final (opcional)
        # shutil.rmtree(tmp_root)
