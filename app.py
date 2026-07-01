import streamlit as st
import pandas as pd
from io import BytesIO

# =============================
# CONFIGURACIÓN
# =============================
st.set_page_config(page_title="Procesador AVE", layout="wide")
st.title("📊 Procesador de Cursos AVE - Versión Auditoría")

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
# NORMALIZACIÓN ROBUSTA
# =============================
def clean_cedula(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"[^\d]", "", regex=True)
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

    df_part["Cedula_clean"] = clean_cedula(df_part["Número de ID"])

    df_part["Estado Cedula"] = df_part["Cedula_clean"].apply(
        lambda x: "VACÍO" if pd.isna(x)
        else "DUPLICADO" if df_part["Cedula_clean"].duplicated(keep=False).loc[df_part["Cedula_clean"] == x].any()
        else "NO DUPLICADO"
    )

    df_part["Nombre Completo"] = (
        df_part["Nombre"].astype(str).str.strip() + " " +
        df_part["Apellido(s)"].astype(str).str.strip()
    )

    df_final = df_part[[
        "Nombre Completo",
        "Número de ID",
        "Cedula_clean",
        "Departamento",
        "Institución",
        "Estado Cedula"
    ]].copy()

    df_final.columns = [
        "Nombre Completo",
        "Cedula",
        "Cedula_clean",
        "Cargo",
        "Seccional",
        "Estado Cedula"
    ]

    # =============================
    # CURSO CON NOTA
    # =============================
    if tipo_curso == "Curso CON nota":

        df_calif = pd.read_excel(file_extra, engine="openpyxl")
        df_calif.columns = df_calif.columns.str.strip()
        df_calif["Cedula_clean"] = clean_cedula(df_calif["Número de ID"])

        df_calif = df_calif.drop_duplicates(subset=["Cedula_clean"], keep="first")

        score_col = "Total del curso (Real)"
        df_calif = df_calif[["Cedula_clean", score_col]].copy()
        df_calif.columns = ["Cedula_clean", "Nota"]
        df_calif["Nota"] = pd.to_numeric(df_calif["Nota"], errors="coerce")

        df_final = df_final.merge(df_calif, on="Cedula_clean", how="left")

        df_final["Aprobo"] = df_final["Nota"].apply(
            lambda x: "No tiene nota" if pd.isna(x)
            else "Sí" if x >= 3.5
            else "No"
        )

        total_aprobados = int((df_final["Aprobo"] == "Sí").sum())
        total_certificados = 0

    # =============================
    # CURSO SIN NOTA (CERTIFICADOS)
    # =============================
    else:

        df_cert = pd.read_excel(file_extra, engine="openpyxl")
        df_cert.columns = df_cert.columns.str.strip()
        df_cert["Cedula_clean"] = clean_cedula(df_cert["Número de ID"])

        # 🔥 SOLO PARA REFERENCIA, NO PARA ALTERAR EL RESULTADO
        df_cert_unique = df_cert.drop_duplicates(subset=["Cedula_clean"], keep="first")

        df_final = df_final.merge(
            df_cert_unique[["Cedula_clean"]],
            on="Cedula_clean",
            how="left",
            indicator=True
        )

        df_final["Aprobo"] = df_final["_merge"].apply(lambda x: "Sí" if x == "both" else "No")
        df_final.drop(columns=["_merge"], inplace=True)

        df_final["Nota"] = ""

        # 🔥 KPI CORRECTO: SOLO BASE ORIGINAL
        total_certificados = df_cert["Cedula_clean"].nunique(dropna=True)
        total_aprobados = int((df_final["Aprobo"] == "Sí").sum())

        # =============================
        # AUDITORÍA DE DIFERENCIA
        # =============================
        no_match = df_final.loc[
            df_final["Aprobo"] == "No",
            ["Cedula", "Nombre Completo"]
        ]

        st.subheader("⚠️ Auditoría")
        st.write(f"Certificados base: {total_certificados}")
        st.write(f"Certificados en resultado (match): {total_aprobados}")
        st.write(f"Diferencia: {total_certificados - total_aprobados}")
        st.dataframe(no_match)

    # =============================
    # KPIs
    # =============================
    st.subheader("📌 Resumen")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Participantes únicos", df_part["Cedula_clean"].nunique(dropna=True))
    c2.metric("Duplicados", int((df_part["Estado Cedula"] == "DUPLICADO").sum()))
    c3.metric("Vacíos", int(df_part["Cedula_clean"].isna().sum()))
    c4.metric("Aprobados", total_aprobados)
    c5.metric("Certificados únicos", total_certificados)

    # =============================
    # OUTPUT
    # =============================
    st.subheader("📄 Resultado final")
    st.dataframe(df_final, use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Resultado")

    st.download_button(
        "📥 Descargar resultado",
        data=output.getvalue(),
        file_name="Resultado_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
