import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

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
# PROCESO
# =============================
if st.button("Procesar"):

    if not file_part or not file_extra:
        st.warning("Debe cargar todos los archivos")
        st.stop()

    # =============================
    # CARGA DATA
    # =============================
    df_part = pd.read_excel(file_part, engine="openpyxl")
    df_extra = pd.read_excel(file_extra, engine="openpyxl")

    df_part.columns = df_part.columns.str.strip()
    df_extra.columns = df_extra.columns.str.strip()

    # =============================
    # DETECTAR CÉDULA
    # =============================
    def find_cedula(df):
        for c in df.columns:
            if "cedula" in c.lower() or "documento" in c.lower():
                return c
        return None

    cedula_part = find_cedula(df_part)
    cedula_extra = find_cedula(df_extra)

    if not cedula_part or not cedula_extra:
        st.error("No se encontró columna de cédula")
        st.stop()

    df_part["_cedula"] = df_part[cedula_part].astype(str).str.strip()
    df_extra["_cedula"] = df_extra[cedula_extra].astype(str).str.strip()

    # =============================
    # VALIDACIÓN CÉDULAS
    # =============================
    mask_vacia = df_part["_cedula"].isin(["", "nan", "None"])
    mask_dup = df_part["_cedula"].duplicated(keep=False) & ~mask_vacia

    cedulas_vacias = df_part.loc[mask_vacia, "_cedula"].tolist()
    cedulas_duplicadas = df_part.loc[mask_dup, "_cedula"].unique().tolist()

    # =============================
    # LIMPIEZA
    # =============================
    df_clean = df_part[~mask_vacia].drop_duplicates(subset="_cedula", keep="first")

    # =============================
    # CRUCE SEGÚN TIPO
    # =============================
    if tipo_curso == "Curso CON nota":

        # buscar columna dinámica de nota
        score_col = None
        for c in df_extra.columns:
            if c.strip() == "Total del curso (Real)":
                score_col = c
                break

        if not score_col:
            st.error('No se encontró "Total del curso (Real)"')
            st.stop()

        df_merge = df_clean.merge(
            df_extra[["_cedula", score_col]],
            on="_cedula",
            how="left"
        )

        df_merge["estado"] = df_merge[score_col].apply(
            lambda x: "CERTIFICADO" if pd.notna(x) and float(x) >= 70
            else ("NO CERTIFICADO" if pd.notna(x) else "SIN NOTA")
        )

        total_cert = int((df_merge["estado"] == "CERTIFICADO").sum())
        total_no_cert = int((df_merge["estado"] == "NO CERTIFICADO").sum())

    else:

        # sin nota → solo validación de existencia
        df_merge = df_clean.merge(
            df_extra[["_cedula"]],
            on="_cedula",
            how="left",
            indicator=True
        )

        df_merge["estado"] = df_merge["_merge"].apply(
            lambda x: "CERTIFICADO" if x == "both" else "NO CERTIFICADO"
        )

        total_cert = int((df_merge["estado"] == "CERTIFICADO").sum())
        total_no_cert = int((df_merge["estado"] == "NO CERTIFICADO").sum())

    # =============================
    # KPIs
    # =============================
    st.subheader("📌 Resumen general")

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Participantes", len(df_part))
    c2.metric("Sin duplicados", len(df_clean))
    c3.metric("Duplicados", len(cedulas_duplicadas))
    c4.metric("Vacíos", len(cedulas_vacias))
    c5.metric("Certificados", total_cert)

    st.markdown("---")

    # =============================
    # REPORTES
    # =============================
    left, right = st.columns(2)

    with left:
        st.subheader("🔴 Cédulas duplicadas")
        st.write(cedulas_duplicadas if cedulas_duplicadas else "No hay duplicados")

    with right:
        st.subheader("⚠️ Cédulas vacías")
        st.write(cedulas_vacias if cedulas_vacias else "No hay vacíos")

    # =============================
    # TABLA
    # =============================
    st.subheader("📄 Datos procesados")

    def color_rows(row):
        if row["_cedula"] in cedulas_duplicadas or row["_cedula"] in cedulas_vacias:
            return ["background-color: #f4cccc"] * len(row)
        return [""] * len(row)

    st.dataframe(df_merge.style.apply(color_rows, axis=1), use_container_width=True)

    # =============================
    # DESCARGA EXCEL
    # =============================
    output = BytesIO()
    df_merge.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        "⬇️ Descargar Excel",
        data=output,
        file_name="resultado_procesado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =============================
    # RESUMEN JSON
    # =============================
    resumen = {
        "total_participantes": len(df_part),
        "sin_duplicados": len(df_clean),
        "duplicados": len(cedulas_duplicadas),
        "vacias": len(cedulas_vacias),
        "certificados": total_cert,
        "no_certificados": total_no_cert,
        "lista_duplicados": cedulas_duplicadas,
        "lista_vacias": cedulas_vacias
    }

    st.download_button(
        "⬇️ Descargar resumen",
        data=pd.DataFrame([resumen]).to_json(orient="records", force_ascii=False),
        file_name="resumen.json",
        mime="application/json"
    )
