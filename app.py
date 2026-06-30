import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Procesador AVE", layout="wide")
st.title("📊 Procesador de Cursos AVE")

tipo_curso = st.radio(
    "Seleccione tipo de curso:",
    ["Curso CON nota", "Curso SIN nota"]
)

file_part = st.file_uploader("📂 Suba archivo PARTICIPANTES", type=["xlsx"])

if tipo_curso == "Curso CON nota":
    file_extra = st.file_uploader("📂 Suba archivo CALIFICACIONES", type=["xlsx"])
else:
    file_extra = st.file_uploader("📂 Suba archivo CERTIFICADOS", type=["xlsx"])


# =============================
# DETECTAR CÉDULA
# =============================
def find_cedula(df):
    for c in df.columns:
        col = str(c).lower().strip()
        if (
            "cedula" in col or
            "cédula" in col or
            "documento" in col or
            "número de id" in col or
            "numero de id" in col or
            "id" in col
        ):
            return c
    return None


# =============================
# PROCESO
# =============================
if st.button("Procesar"):

    if not file_part or not file_extra:
        st.warning("Debe cargar ambos archivos")
        st.stop()

    df_part = pd.read_excel(file_part, engine="openpyxl")
    df_extra = pd.read_excel(file_extra, engine="openpyxl")

    df_part.columns = df_part.columns.str.strip()
    df_extra.columns = df_extra.columns.str.strip()

    cedula_part = find_cedula(df_part)
    cedula_extra = find_cedula(df_extra)

    if not cedula_part or not cedula_extra:
        st.error("No se encontró columna de ID")
        st.stop()

    # =============================
    # NORMALIZACIÓN (CLAVE)
    # =============================
    df_part["Cedula"] = (
        df_part[cedula_part]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    df_extra["Cedula"] = (
        df_extra[cedula_extra]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    # =============================
    # VACÍOS (ROBUSTO)
    # =============================
    mask_vacia = df_part["Cedula"].isna() | df_part["Cedula"].isin(["", "nan", "None", "NaN"])

    # =============================
    # DUPLICADOS (NO SE BORRAN)
    # =============================
    mask_dup = df_part["Cedula"].duplicated(keep=False) & ~mask_vacia

    cedulas_vacias = df_part.loc[mask_vacia, "Cedula"].unique().tolist()
    cedulas_duplicadas = df_part.loc[mask_dup, "Cedula"].unique().tolist()

    # =============================
    # BASE (SIN BORRAR DUPLICADOS)
    # =============================
    df_base = df_part.copy()

    df_base["Nombre Completo"] = (
        df_base["Nombre"].astype(str).str.strip() + " " +
        df_base["Apellido(s)"].astype(str).str.strip()
    )

    df_final = df_base[[
        "Nombre Completo",
        "Cedula",
        "Departamento",
        "Institución"
    ]].copy()

    df_final.columns = ["Nombre Completo", "Cedula", "Cargo", "Seccional"]

    # =============================
    # CURSO CON NOTA
    # =============================
    if tipo_curso == "Curso CON nota":

        score_col = None
        for c in df_extra.columns:
            if str(c).strip() == "Total del curso (Real)":
                score_col = c
                break

        if not score_col:
            st.error("No se encontró 'Total del curso (Real)'")
            st.stop()

        df_extra["Cedula"] = (
            df_extra["Cedula"].astype(str).str.strip()
        )

        df_calif = df_extra[["Cedula", score_col]].copy()
        df_calif.columns = ["Cedula", "Nota"]

        # ✅ IMPORTANTE: limpieza de nota
        df_calif["Nota"] = pd.to_numeric(df_calif["Nota"], errors="coerce")

        # =============================
        # CRUCE CORRECTO
        # =============================
        df_final = df_final.merge(
            df_calif,
            on="Cedula",
            how="left"
        )

        # =============================
        # ESTADO CORRECTO
        # =============================
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

        df_extra["Cedula"] = df_extra["Cedula"].astype(str).str.strip()

        df_extra["Aprobo_flag"] = 1

        df_final = df_final.merge(
            df_extra[["Cedula", "Aprobo_flag"]],
            on="Cedula",
            how="left"
        )

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
    c2.metric("Sin duplicados", len(df_part))  # 👈 NO se elimina
    c3.metric("Duplicados", len(cedulas_duplicadas))
    c4.metric("Vacíos", len(cedulas_vacias))
    c5.metric("Aprobados", total_aprobados)

    # =============================
    # REPORTES
    # =============================
    st.subheader("🔴 Duplicados")
    st.write(f"Total: {len(cedulas_duplicadas)}")
    st.write(cedulas_duplicadas)

    st.subheader("⚠️ Vacíos")
    st.write(f"Total: {len(cedulas_vacias)}")
    st.write(cedulas_vacias)

    # =============================
    # COLORACION
    # =============================
    def color_rows(row):
        if row["Cedula"] in cedulas_duplicadas or row["Cedula"] in cedulas_vacias:
            return ["background-color: #f4cccc"] * len(row)
        return [""] * len(row)

    st.subheader("📄 Resultado")
    st.dataframe(df_final.style.apply(color_rows, axis=1), use_container_width=True)

    # =============================
    # DESCARGA
    # =============================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Resultado")

    st.download_button(
        "📥 Descargar resultado",
        data=output.getvalue(),
        file_name="Resultado_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
