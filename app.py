import streamlit as st
import pandas as pd
from io import BytesIO

# =============================
# CONFIGURACIÓN
# =============================
st.set_page_config(page_title="Procesador AVE", layout="wide")
st.title("📊 Procesador de Cursos AVE")

# =============================
# TIPO DE CURSO
# =============================
tipo_curso = st.radio(
    "Seleccione tipo de curso:",
    ["Curso CON nota", "Curso SIN nota"]
)

# =============================
# ARCHIVOS
# =============================
file_part = st.file_uploader("📂 Suba archivo PARTICIPANTES", type=["xlsx"])

if tipo_curso == "Curso CON nota":
    file_extra = st.file_uploader("📂 Suba archivo CALIFICACIONES", type=["xlsx"])
else:
    file_extra = st.file_uploader("📂 Suba archivo CERTIFICADOS", type=["xlsx"])

# =============================
# NORMALIZACIÓN (CLAVE)
# =============================
def clean_cedula(series):
    return (
        series.astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"[^0-9]", "", regex=True)
        .str.strip()
        .replace({"nan": None, "None": None, "": None})
    )

# =============================
# PROCESO
# =============================
if st.button("Procesar"):

    if not file_part or not file_extra:
        st.warning("Debe cargar ambos archivos")
        st.stop()

    # =============================
    # PARTICIPANTES
    # =============================
    df_part = pd.read_excel(file_part, engine="openpyxl")
    df_part.columns = df_part.columns.str.strip()

    st.subheader("📊 Participantes")
    st.dataframe(df_part.head())

    # ✅ CÉDULA LIMPIA
    df_part["Cedula"] = clean_cedula(df_part["Número de ID"])

    # ✅ VACÍOS
    mask_vacia = df_part["Cedula"].isna()

    # ✅ DUPLICADOS REALES
    mask_dup = df_part["Cedula"].notna() & df_part["Cedula"].duplicated(keep=False)

    cedulas_duplicadas = df_part.loc[mask_dup, "Cedula"].unique().tolist()
    cedulas_vacias = df_part.loc[mask_vacia, "Cedula"].tolist()

    # =============================
    # NOMBRE COMPLETO
    # =============================
    df_part["Nombre Completo"] = (
        df_part["Nombre"].astype(str).str.strip() + " " +
        df_part["Apellido(s)"].astype(str).str.strip()
    )

    # =============================
    # BASE FINAL
    # =============================
    df_final = df_part[[
        "Nombre Completo",
        "Cedula",
        "Departamento",
        "Institución"
    ]].copy()

    df_final.columns = ["Nombre Completo", "Cedula", "Cargo", "Seccional"]

    # =============================
    # DUPLICADO (COL PILAR)
    # =============================
    def label_dup(x):
        if pd.isna(x):
            return "VACÍO"
        elif x in cedulas_duplicadas:
            return "DUPLICADO"
        else:
            return "NO DUPLICADO"

    df_final["Estado Cedula"] = df_final["Cedula"].apply(label_dup)

    # =============================
    # CURSO CON NOTA
    # =============================
    if tipo_curso == "Curso CON nota":

        df_calif = pd.read_excel(file_extra, engine="openpyxl")
        df_calif.columns = df_calif.columns.str.strip()

        st.subheader("📄 Calificaciones")
        st.dataframe(df_calif.head())

        df_calif["Cedula"] = clean_cedula(df_calif["Número de ID"])

        # ✅ BUSCAR NOTA
        score_col = "Total del curso (Real)"
        if score_col not in df_calif.columns:
            st.error("No se encontró la columna 'Total del curso (Real)'")
            st.stop()

        df_calif = df_calif[["Cedula", score_col]].copy()
        df_calif.columns = ["Cedula", "Nota"]

        df_calif["Nota"] = pd.to_numeric(df_calif["Nota"], errors="coerce")

        # ✅ CRUCE
        df_final = df_final.merge(df_calif, on="Cedula", how="left")

        df_final["Aprobo"] = df_final["Nota"].apply(
            lambda x: "No tiene nota" if pd.isna(x)
            else "Sí" if x >= 3.5
            else "No"
        )

        total_aprobados = int((df_final["Aprobo"] == "Sí").sum())

    # =============================
    # CURSO SIN NOTA
    # =============================
    else:

        df_cert = pd.read_excel(file_extra, engine="openpyxl")
        df_cert.columns = df_cert.columns.str.strip()

        st.subheader("📄 Certificados")
        st.dataframe(df_cert.head())

        df_cert["Cedula"] = clean_cedula(df_cert["Número de ID"])

        df_cert["Aprobo_flag"] = 1

        df_final = df_final.merge(
            df_cert[["Cedula", "Aprobo_flag"]],
            on="Cedula",
            how="left"
        )

        df_final["Nota"] = ""
        df_final["Aprobo"] = df_final["Aprobo_flag"].apply(
            lambda x: "Sí" if pd.notnull(x) else "No"
        )

        df_final.drop(columns=["Aprobo_flag"], inplace=True)

        total_aprobados = int((df_final["Aprobo"] == "Sí").sum())

    # =============================
    # KPIs
    # =============================
    st.subheader("📌 Resumen")

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Participantes", len(df_part))
    c2.metric("Duplicados", len(cedulas_duplicadas))
    c3.metric("Vacíos", int(mask_vacia.sum()))
    c4.metric("Aprobados", total_aprobados)
    c5.metric("Registros finales", len(df_final))

    # =============================
    # TABLA FINAL (VISUAL)
    # =============================
    st.subheader("📄 Resultado final")
    st.dataframe(df_final, use_container_width=True)

    # =============================
    # EXPORTACIÓN EXCEL (SIN ESTILO, CON COL)
    # =============================
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Resultado")

        # =============================
        # OPCIONAL: COLOR EN EXCEL
        # =============================
        workbook = writer.book
        worksheet = writer.sheets["Resultado"]

        format_dup = workbook.add_format({"bg_color": "#f4cccc"})

        for i, val in enumerate(df_final["Estado Cedula"], start=1):
            if val == "DUPLICADO":
                worksheet.set_row(i, None, format_dup)

    st.download_button(
        "📥 Descargar resultado",
        data=output.getvalue(),
        file_name="Resultado_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
