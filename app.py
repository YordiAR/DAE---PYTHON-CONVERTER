import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

# =============================
# CONFIGURACIÓN DE LA PÁGINA
# =============================
st.set_page_config(page_title="Validación de Participantes", layout="wide")

st.title("📊 Validación de Participantes y Calificaciones")

# =============================
# SUBIDA DE ARCHIVOS
# =============================
col1, col2 = st.columns(2)

with col1:
    participantes_file = st.file_uploader("📂 Cargar archivo de Participantes", type=["xlsx"])

with col2:
    calificaciones_file = st.file_uploader("📂 Cargar archivo de Calificaciones", type=["xlsx"])

if not participantes_file or not calificaciones_file:
    st.info("Carga ambos archivos para continuar")
    st.stop()

# =============================
# CARGA DATAFRAME
# =============================
df_part = pd.read_excel(participantes_file, engine="openpyxl")
df_cal = pd.read_excel(calificaciones_file, engine="openpyxl")

df_part.columns = df_part.columns.str.strip()
df_cal.columns = df_cal.columns.str.strip()

# =============================
# DETECTAR CÉDULA
# =============================
def find_cedula(df):
    for c in df.columns:
        if "cedula" in c.lower() or "documento" in c.lower():
            return c
    return None

cedula_part = find_cedula(df_part)
cedula_cal = find_cedula(df_cal)

if not cedula_part or not cedula_cal:
    st.error("No se encontró columna de cédula en uno de los archivos")
    st.stop()

df_part["_cedula"] = df_part[cedula_part].astype(str).str.strip()
df_cal["_cedula"] = df_cal[cedula_cal].astype(str).str.strip()

# =============================
# VALIDACIÓN DE DATOS
# =============================
mask_vacia = df_part["_cedula"].isin(["", "nan", "None"])
mask_dup = df_part["_cedula"].duplicated(keep=False) & ~mask_vacia

cedulas_vacias = df_part.loc[mask_vacia, "_cedula"].tolist()
cedulas_duplicadas = df_part.loc[mask_dup, "_cedula"].unique().tolist()

df_clean = df_part[~mask_vacia].drop_duplicates(subset="_cedula", keep="first")

# =============================
# BUSCAR COLUMNA DE NOTA
# =============================
score_col = None
for c in df_cal.columns:
    if c.strip() == "Total del curso (Real)":
        score_col = c
        break

if not score_col:
    st.error('No se encontró la columna "Total del curso (Real)"')
    st.stop()

# =============================
# CRUCE
# =============================
df_merge = df_clean.merge(
    df_cal[["_cedula", score_col]],
    on="_cedula",
    how="left"
)

# =============================
# ESTADO
# =============================
df_merge["estado"] = df_merge[score_col].apply(
    lambda x: "CERTIFICADO" if pd.notna(x) and float(x) >= 70
    else ("NO CERTIFICADO" if pd.notna(x) else "SIN NOTA")
)

# =============================
# KPIs
# =============================
st.subheader("📌 Resumen general")

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Participantes", len(df_part))
k2.metric("Sin duplicados", len(df_clean))
k3.metric("Duplicados", len(cedulas_duplicadas))
k4.metric("Vacíos", len(cedulas_vacias))
k5.metric("Certificados", int((df_merge["estado"] == "CERTIFICADO").sum()))

st.markdown("---")

# =============================
# REPORTES
# =============================
colA, colB = st.columns(2)

with colA:
    st.subheader("🔴 Cédulas duplicadas")
    st.write(cedulas_duplicadas if cedulas_duplicadas else "No hay duplicados")

with colB:
    st.subheader("⚠️ Cédulas vacías")
    st.write(cedulas_vacias if cedulas_vacias else "No hay vacíos")

# =============================
# TABLA INTERACTIVA
# =============================
st.subheader("📄 Vista de datos procesados")

def color_rows(row):
    if row["_cedula"] in cedulas_duplicadas or row["_cedula"] in cedulas_vacias:
        return ["background-color: #f4cccc"] * len(row)
    return [""] * len(row)

st.dataframe(df_merge.style.apply(color_rows, axis=1), use_container_width=True)

# =============================
# DESCARGA EXCEL
# =============================
output = io.BytesIO()
df_merge.to_excel(output, index=False, engine="openpyxl")
output.seek(0)

st.download_button(
    label="⬇️ Descargar resultado en Excel",
    data=output,
    file_name="resultado_cruzado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =============================
# DESCARGA RESUMEN
# =============================
resumen = {
    "total_participantes": len(df_part),
    "sin_duplicados": len(df_clean),
    "duplicados": len(cedulas_duplicadas),
    "vacias": len(cedulas_vacias),
    "certificados": int((df_merge["estado"] == "CERTIFICADO").sum()),
    "no_certificados": int((df_merge["estado"] == "NO CERTIFICADO").sum()),
    "lista_duplicados": cedulas_duplicadas,
    "lista_vacias": cedulas_vacias
}

st.download_button(
    label="⬇️ Descargar resumen JSON",
    data=pd.DataFrame([resumen]).to_json(orient="records", force_ascii=False),
    file_name="resumen.json",
    mime="application/json"
)
