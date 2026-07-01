import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Procesador AVE - Corregido", layout="wide")
st.title("📊 Procesador de Cursos AVE (Versión Corregida)")

# =====================================
# 🔧 NORMALIZACIÓN DE TEXTO
# =====================================
def normalize_text(series):
    return (
        series.astype(str)
        .str.strip()
        .str.upper()
        .str.replace("Á","A", regex=False)
        .str.replace("É","E", regex=False)
        .str.replace("Í","I", regex=False)
        .str.replace("Ó","O", regex=False)
        .str.replace("Ú","U", regex=False)
        .str.replace("Ñ","N", regex=False)
        .str.replace(r"\s+", " ", regex=True)
    )

# =====================================
# 🔎 DETECCIÓN INTELIGENTE DE CÉDULA
# =====================================
def find_id_column(df):
    posibles = [
        "cedula", "cédula",
        "numero de id", "número de id",
        "documento", "documento de identidad",
        "identificacion", "identificación",
        "id", "no identificacion"
    ]

    candidatos = []
    for c in df.columns:
        cl = str(c).lower()
        if any(k in cl for k in posibles):
            candidatos.append(c)

    # elegir la columna que más parece ID (más números válidos)
    for c in candidatos:
        sample = df[c].astype(str).str.replace(r"\D", "", regex=True)
        if sample.str.len().median() >= 7:
            return c

    return None

# =====================================
# 🔎 DETECTAR OTRAS COLUMNAS
# =====================================
def find_column(df, keywords):
    for k in keywords:
        col = next((c for c in df.columns if k in c.lower()), None)
        if col:
            return col
    return None

# =====================================
# 🟢 TIPO DE CURSO
# =====================================
tipo_curso = st.radio("Seleccione tipo de curso:", ["Curso CON nota", "Curso SIN nota"])

file_part = st.file_uploader("📂 Participantes", type=["xlsx"])

if tipo_curso == "Curso CON nota":
    file_extra = st.file_uploader("📂 Calificaciones", type=["xlsx"])
else:
    file_extra = st.file_uploader("📂 Aprobados", type=["xlsx"])

# =====================================
# 🚀 PROCESO
# =====================================
if st.button("Procesar"):

    if not file_part or not file_extra:
        st.warning("⚠️ Debes subir ambos archivos")
        st.stop()

    try:
        # =====================================
        # 📊 PARTICIPANTES
        # =====================================
        df_part = pd.read_excel(file_part)
        df_part.columns = df_part.columns.str.strip()

        col_id_part = find_id_column(df_part)
        col_nombre = find_column(df_part, ["nombre"])
        col_apellido = find_column(df_part, ["apellido"])

        if not col_id_part:
            st.error("❌ No se encontró columna válida de cédula en participantes")
            st.stop()

        # =====================================
        # 🔥 LIMPIEZA REAL DE CÉDULA (CRÍTICO)
        # =====================================
        df_part["Cedula"] = (
            df_part[col_id_part]
            .astype(str)
            .str.replace(r"\D", "", regex=True)
        )

        # eliminar inválidos
        df_part = df_part[
            df_part["Cedula"].notna() &
            (df_part["Cedula"] != "") &
            (df_part["Cedula"].str.len() >= 7)
        ].copy()

        # eliminar duplicados reales
        df_part = df_part.drop_duplicates(subset=["Cedula"])

        # =====================================
        # 👤 NOMBRES
        # =====================================
        if col_nombre and col_apellido:
            df_part["Nombres y Apellidos"] = (
                normalize_text(df_part[col_nombre]) + " " +
                normalize_text(df_part[col_apellido])
            )
        elif col_nombre:
            df_part["Nombres y Apellidos"] = normalize_text(df_part[col_nombre])
        else:
            df_part["Nombres y Apellidos"] = ""

        # =====================================
        # 🧠 BASE FINAL
        # =====================================
        base = df_part[[
            "Cedula",
            "Nombres y Apellidos",
            "Departamento",
            "Institución"
        ]].copy()

        base.columns = ["Cedula", "Nombres y Apellidos", "Cargo", "Dependencia"]

        # =====================================
        # 🟢 CON NOTA
        # =====================================
        if tipo_curso == "Curso CON nota":

            df_calif = pd.read_excel(file_extra)
            df_calif.columns = df_calif.columns.str.strip()

            col_id_calif = find_id_column(df_calif)
            col_nota = find_column(df_calif, ["total del curso", "nota", "calificacion"])

            if not col_id_calif or not col_nota:
                st.error("❌ Calificaciones: estructura inválida")
                st.stop()

            df_calif["Cedula"] = (
                df_calif[col_id_calif]
                .astype(str)
                .str.replace(r"\D", "", regex=True)
            )

            df_calif["Nota"] = pd.to_numeric(df_calif[col_nota], errors="coerce")

            # quitar duplicados (mejor nota)
            df_calif = df_calif.sort_values("Nota", ascending=False).drop_duplicates("Cedula")

            resultado = base.merge(
                df_calif[["Cedula", "Nota"]],
                on="Cedula",
                how="left"
            )

            resultado["Aprobo"] = resultado["Nota"].apply(
                lambda x: "Sí" if pd.notnull(x) and x >= 3.5 else "No"
            )

        # =====================================
        # 🔵 SIN NOTA (CORREGIDO)
        # =====================================
        else:

            df_aprob = pd.read_excel(file_extra)
            df_aprob.columns = df_aprob.columns.str.strip()

            col_id_aprob = find_id_column(df_aprob)

            if not col_id_aprob:
                st.error("❌ No se encontró columna válida de cédula en aprobados")
                st.stop()

            df_aprob["Cedula"] = (
                df_aprob[col_id_aprob]
                .astype(str)
                .str.replace(r"\D", "", regex=True)
            )

            # eliminar inválidos
            df_aprob = df_aprob[
                df_aprob["Cedula"].notna() &
                (df_aprob["Cedula"] != "") &
                (df_aprob["Cedula"].str.len() >= 7)
            ].copy()

            # eliminar duplicados reales
            df_aprob = df_aprob.drop_duplicates(subset=["Cedula"])

            resultado = base.merge(
                df_aprob[["Cedula"]],
                on="Cedula",
                how="left",
                indicator=True
            )

            resultado["Nota"] = ""
            resultado["Aprobo"] = resultado["_merge"].apply(
                lambda x: "Sí" if x == "both" else "No"
            )
            resultado.drop(columns=["_merge"], inplace=True)

        # =====================================
        # ⚠️ IRREGULARIDADES
        # =====================================
        resultado["Irregularidades"] = ""

        # sin ID
        resultado.loc[resultado["Cedula"] == "", "Irregularidades"] += "SIN ID; "

        # no encontrado
        resultado.loc[resultado["Aprobo"] == "No", "Irregularidades"] += "NO ENCONTRADO; "

        # =====================================
        # 📊 CONTROL
        # =====================================
        st.subheader("📊 Control")

        col1, col2, col3 = st.columns(3)
        col1.metric("Participantes", len(df_part))
        col2.metric("Aprobados", len(df_aprob))
        col3.metric("Resultado", len(resultado))

        st.write("🔎 Cédulas únicas participantes:", df_part["Cedula"].nunique())
        st.write("🔎 Cédulas únicas aprobados:", df_aprob["Cedula"].nunique())

        # =====================================
        # 📥 EXPORTAR
        # =====================================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            resultado.to_excel(writer, index=False, sheet_name="Resultado")

        output.seek(0)

        wb = load_workbook(output)
        ws = wb["Resultado"]

        rojo = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        amarillo = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

        col_irreg = resultado.columns.get_loc("Irregularidades")

        for row in ws.iter_rows(min_row=2):
            val = str(row[col_irreg].value)

            if "SIN ID" in val:
                for cell in row:
                    cell.fill = amarillo
            elif "NO ENCONTRADO" in val:
                for cell in row:
                    cell.fill = rojo

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.success("✅ Procesado correctamente")
        st.dataframe(resultado)

        st.download_button(
            "📥 Descargar resultado",
            data=final_output.getvalue(),
            file_name="Resultado_Cursos_AVE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
