import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill

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
# DETECTOR DE COLUMNA ID (CORREGIDO)
# =============================
def find_cedula(df):
    posibles = [
        "cedula", "cédula",
        "documento",
        "número de id", "numero de id",
        "id", "identificacion", "identificación"
    ]

    for c in df.columns:
        col = str(c).lower().strip()

        # match exacto o parcial
        if col in posibles or any(p in col for p in posibles):
            return c

    return None

# =============================
# PROCESO PRINCIPAL
# =============================
if st.button("Procesar"):

    if not file_part or not file_extra:
        st.warning("Debe cargar ambos archivos")
        st.stop()

    # =============================
    # CARGA
    # =============================
    df_part = pd.read_excel(file_part, engine="openpyxl")
    df_extra = pd.read_excel(file_extra, engine="openpyxl")

    df_part.columns = df_part.columns.str.strip()
    df_extra.columns = df_extra.columns.str.strip()

    # =============================
    # DETECTAR ID
    # =============================
    cedula_part = find_cedula(df_part)
    cedula_extra = find_cedula(df_extra)

    if not cedula_part:
        st.error("No se encontró columna de ID en PARTICIPANTES")
        st.write("Columnas disponibles:", df_part.columns.tolist())
        st.stop()

    if not cedula_extra:
        st.error("No se encontró columna de ID en CALIFICACIONES/CERTIFICADOS")
        st.write("Columnas disponibles:", df_extra.columns.tolist())
        st.stop()

    # =============================
    # NORMALIZAR
    # =============================
    df_part["_id"] = df_part[cedula_part].astype(str).str.strip()
    df_extra["_id"] = df_extra[cedula_extra].astype(str).str.strip()

    # =============================
    # VACÍOS Y DUPLICADOS
    # =============================
    mask_vacia = df_part["_id"].isin(["", "nan", "None"])
    mask_dup = df_part["_id"].duplicated(keep=False) & ~mask_vacia

    cedulas_vacias = df_part.loc[mask_vacia, "_id"].tolist()
    cedulas_duplicadas = df_part.loc[mask_dup, "_id"].unique().tolist()

    # =============================
    # LIMPIEZA
    # =============================
    df_clean = df_part[~mask_vacia].drop_duplicates(subset="_id", keep="first")

    # =============================
    # CRUCE
    # =============================
    if tipo_curso == "Curso CON nota":

        # buscar nota sin depender de posición
        score_col = None
        for c in df_extra.columns:
            if str(c).strip() == "Total del curso (Real)":
                score_col = c
                break

        if not score_col:
            st.error("No se encontró 'Total del curso (Real)'")
            st.write("Columnas disponibles:", df_extra.columns.tolist())
            st.stop()

        df_merge = df_clean.merge(
            df_extra[["_id", score_col]],
            on="_id",
            how="left"
        )

        df_merge["estado"] = df_merge[score_col].apply(
            lambda x: "CERTIFICADO" if pd.notna(x) and float(x) >= 70
            else ("NO CERTIFICADO" if pd.notna(x) else "SIN NOTA")
        )

    else:

        df_merge = df_clean.merge(
            df_extra[["_id"]],
            on="_id",
            how="left",
            indicator=True
        )

        df_merge["estado"] = df_merge["_merge"].apply(
            lambda x: "CERTIFICADO" if x == "both" else "NO CERTIFICADO"
        )

    # =============================
    # KPIs
    # =============================
    st.subheader("📌 Resumen")

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Participantes", len(df_part))
    c2.metric("Sin duplicados", len(df_clean))
    c3.metric("Duplicados", len(cedulas_duplicadas))
    c4.metric("Vacíos", len(cedulas_vacias))
    c5.metric("Certificados", int((df_merge["estado"] == "CERTIFICADO").sum()))

    st.markdown("---")

    # =============================
    # REPORTES
    # =============================
    colA, colB = st.columns(2)

    with colA:
        st.subheader("🔴 Duplicados")
        st.write(cedulas_duplicadas if cedulas_duplicadas else "Sin duplicados")

    with colB:
        st.subheader("⚠️ Vacíos")
        st.write(cedulas_vacias if cedulas_vacias else "Sin vacíos")

    # =============================
    # TABLA CON MARCADO
    # =============================
    def color_rows(row):
        if row["_id"] in cedulas_duplicadas or row["_id"] in cedulas_vacias:
            return ["background-color: #f4cccc"] * len(row)
        return [""] * len(row)

    st.subheader("📄 Datos procesados")
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
