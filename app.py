import streamlit as st
import pandas as pd
from io import BytesIO

# =============================
# CONFIGURACIÓN
# =============================
st.set_page_config(page_title="Procesador AVE", layout="wide")
st.title("📊 Procesador de Cursos AVE")

# =============================
# SELECCIÓN DE TIPO
# =============================
tipo_curso = st.radio(
    "Seleccione tipo de curso:",
    ["Curso CON nota", "Curso SIN nota"]
)

# =============================
# CARGA DE ARCHIVOS
# =============================
file_part = st.file_uploader("📂 Suba archivo PARTICIPANTES", type=["xlsx"])

if tipo_curso == "Curso CON nota":
    file_extra = st.file_uploader("📂 Suba archivo CALIFICACIONES", type=["xlsx"])
else:
    file_extra = st.file_uploader("📂 Suba archivo CERTIFICADOS", type=["xlsx"])

# =============================
# FUNCIÓN DE NORMALIZACIÓN (CLAVE)
# =============================
def normalize_id(series):
    return (
        series.astype(str)
        .str.replace(r"\.0$", "", regex=True)
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

    st.subheader("📊 Participantes (vista)")
    st.dataframe(df_part.head())

    # ✅ CÉDULA REAL
    df_part["Cedula"] = normalize_id(df_part["Número de ID"])

    # ✅ VACÍOS (REAL)
    mask_vacia = df_part["Cedula"].isna()

    # ✅ DUPLICADOS (NO SE ELIMINAN)
    mask_dup = df_part["Cedula"].notna() & df_part["Cedula"].duplicated(keep=False)

    cedulas_vacias = df_part.loc[mask_vacia, "Cedula"].tolist()
    cedulas_duplicadas = df_part.loc[mask_dup, "Cedula"].unique().tolist()

    # =============================
    # NOMBRE COMPLETO
    # =============================
    df_part["Nombre Completo"] = (
        df_part["Nombre"].astype(str).str.strip() + " " +
        df_part["Apellido(s)"].astype(str).str.strip()
    )

    # =============================
    # BASE FINAL (SIN BORRAR DUPLICADOS)
    # =============================
    df_final = df_part[[
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

        df_calif = pd.read_excel(file_extra, engine="openpyxl")
        df_calif.columns = df_calif.columns.str.strip()

        st.subheader("📄 Calificaciones (vista)")
        st.dataframe(df_calif.head())

        df_calif["Cedula"] = normalize_id(df_calif["Número de ID"])

        # ✅ BUSCAR NOTA DINÁMICA
        score_col = None
        for c in df_calif.columns:
            if str(c).strip() == "Total del curso (Real)":
                score_col = c
                break

        if not score_col:
            st.error("No se encontró la columna 'Total del curso (Real)'")
            st.stop()

        df_calif = df_calif[["Cedula", score_col]].copy()
        df_calif.columns = ["Cedula", "Nota"]

        # ✅ CONVERSIÓN SEGURA
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
        # RESULTADO
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

        df_cert = pd.read_excel(file_extra, engine="openpyxl")
        df_cert.columns = df_cert.columns.str.strip()

        st.subheader("📄 Certificados (vista)")
        st.dataframe(df_cert.head())

        df_cert["Cedula"] = normalize_id(df_cert["Número de ID"])

        # ✅ MARCADOR DE EXISTENCIA
        df_cert["Aprobo_flag"] = 1

        # =============================
        # CRUCE CORRECTO (SIN ELIMINAR)
        # =============================
        df_final = df_final.merge(
            df_cert[["Cedula", "Aprobo_flag"]],
            on="Cedula",
            how="left"
        )

        # =============================
        # RESULTADO
        # =============================
        df_final["Nota"] = ""

        df_final["Aprobo"] = df_final["Aprobo_flag"].apply(
            lambda x: "Sí" if pd.notnull(x) else "No"
        )

        df_final.drop(columns=["Aprobo_flag"], inplace=True)

        total_aprobados = int((df_final["Aprobo"] == "Sí").sum())

    # =============================
    # KPIs
    # =============================
    st.subheader("📌 Resumen general")

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Participantes", len(df_part))
    c2.metric("Base sin duplicar", len(df_part))
    c3.metric("Duplicados", len(cedulas_duplicadas))
    c4.metric("Vacíos", int(mask_vacia.sum()))
    c5.metric("Aprobados", total_aprobados)

    # =============================
    # REPORTES
    # =============================
    colA, colB = st.columns(2)

    with colA:
        st.subheader("🔴 Cédulas duplicadas")
        st.write(f"Total: {len(cedulas_duplicadas)}")
        st.write(cedulas_duplicadas if cedulas_duplicadas else "No hay duplicados")

    with colB:
        st.subheader("⚠️ Cédulas vacías")
        st.write(f"Total: {int(mask_vacia.sum())}")
        st.write(cedulas_vacias if cedulas_vacias else "No hay vacíos")

    # =============================
    # 🎨 MARCADO EN ROJO (DUPLICADOS + VACÍOS)
    # =============================
    def color_rows(row):
        if pd.isna(row["Cedula"]) or row["Cedula"] in cedulas_duplicadas:
            return ["background-color: #f4cccc"] * len(row)
        return [""] * len(row)

    st.subheader("📄 Resultado final")
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
