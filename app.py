import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

st.set_page_config(page_title="Procesador AVE", layout="wide")
st.title("📊 Procesador de Cursos AVE")

# =====================================
# 🔘 TIPO DE CURSO
# =====================================
tipo_curso = st.radio(
    "Seleccione tipo de curso:",
    ["Curso CON nota", "Curso SIN nota"]
)

# =====================================
# 📂 ARCHIVOS
# =====================================
file_part = st.file_uploader("📂 Participantes", type=["xlsx"])

if tipo_curso == "Curso CON nota":
    file_extra = st.file_uploader("📂 Calificaciones", type=["xlsx"])
else:
    file_extra = st.file_uploader("📂 Aprobados", type=["xlsx"])

# =====================================
# 🚀 PROCESAR
# =====================================
if st.button("Procesar"):

    if file_part is None or file_extra is None:
        st.warning("⚠️ Debes subir ambos archivos")
        st.stop()

    # =====================================
    # 📊 PARTICIPANTES
    # =====================================
    df_part = pd.read_excel(file_part)
    df_part.columns = df_part.columns.str.strip()

    # ✅ normalizar
    df_part["Cedula"] = df_part["Número de ID"].astype(str).str.strip()
    df_part["Nombres"] = df_part["Nombre"].astype(str).str.strip() + " " + df_part["Apellido(s)"].astype(str).str.strip()

    # =====================================
    # 🔍 VALIDACIONES PARTICIPANTES
    # =====================================
    df_part["sin_id"] = df_part["Cedula"].isna() | (df_part["Cedula"] == "") | (df_part["Cedula"] == "nan")
    df_part["duplicado_id"] = df_part.duplicated(subset=["Cedula"], keep=False)

    # =====================================
    # 🧠 BASE FINAL
    # =====================================
    base = df_part[[
        "Nombres",
        "Cedula",
        "Departamento",
        "Institución"
    ]].copy()

    base.columns = ["Nombres y Apellidos", "Cedula", "Cargo", "Dependencia"]

    # =====================================
    # 🟢 CON NOTA
    # =====================================
    if tipo_curso == "Curso CON nota":

        df_calif = pd.read_excel(file_extra)
        df_calif.columns = df_calif.columns.str.strip()

        df_calif["Cedula"] = df_calif["Número de ID"].astype(str).str.strip()
        df_calif["Nombres"] = df_calif["Nombre"].astype(str).str.strip() + " " + df_calif["Apellido(s)"].astype(str).str.strip()

        df_calif["Nota"] = pd.to_numeric(df_calif["Total del curso (Real)"], errors="coerce")

        # 🔄 CRUCE DOBLE (ID + Nombres)
        resultado = base.merge(
            df_calif[["Cedula", "Nombres", "Nota"]],
            on=["Cedula", "Nombres"],
            how="left"
        )

        resultado["Aprobo"] = resultado["Nota"].apply(
            lambda x: "Sí" if pd.notnull(x) and x >= 3.5 else "No"
        )

    # =====================================
    # 🔵 SIN NOTA
    # =====================================
    else:


df_aprob = pd.read_excel(file_extra)
df_aprob.columns = df_aprob.columns.str.strip()

# 🔎 detectar columnas
col_id = next((c for c in df_aprob.columns if "id" in c.lower()), None)

if not col_id:
    st.error("❌ No se encontró columna Número de ID en aprobados")
    st.stop()

df_aprob["Cedula"] = df_aprob[col_id].astype(str).str.strip()
df_aprob["Aprobo_flag"] = 1

# ✅ cruce SOLO por cédula (más estable)
resultado = base.merge(
    df_aprob[["Cedula", "Aprobo_flag"]],
    on="Cedula",
    how="left"
)

resultado["Nota"] = ""
resultado["Aprobo"] = resultado["Aprobo_flag"].apply(
    lambda x: "Sí" if pd.notnull(x) else "No"
)

resultado.drop(columns=["Aprobo_flag"], inplace=True)

        # 🔄 CRUCE DOBLE
        resultado = base.merge(
            df_aprob[["Cedula", "Nombres", "Aprobo_flag"]],
            on=["Cedula", "Nombres"],
            how="left"
        )

        resultado["Nota"] = ""
        resultado["Aprobo"] = resultado["Aprobo_flag"].apply(lambda x: "Sí" if pd.notnull(x) else "No")

    # =====================================
    # ⚠️ IRREGULARIDADES
    # =====================================

    resultado["Irregularidades"] = ""

    # 🚨 sin cédula
    resultado.loc[base["Cedula"].isin(["", "nan", None]), "Irregularidades"] += "SIN ID; "

    # 🚨 duplicados
    dup_ids = base[base["Cedula"].duplicated(keep=False)]["Cedula"].tolist()
    resultado.loc[resultado["Cedula"].isin(dup_ids), "Irregularidades"] += "CEDULA DUPLICADA; "

    # 🚨 no encontrado
    resultado.loc[resultado["Aprobo"] == "No", "Irregularidades"] += "NO ENCONTRADO; "

    # =====================================
    # 📊 RESUMEN
    # =====================================
    st.subheader("📊 Resumen")
    st.write("Total participantes:", len(df_part))
    st.write("Sin ID:", df_part["sin_id"].sum())
    st.write("Duplicados:", df_part["duplicado_id"].sum())
    st.write("Aprobados:", (resultado["Aprobo"] == "Sí").sum())

    st.write("IDs sin ID:", df_part[df_part["sin_id"]]["Cedula"].tolist())
    st.write("IDs duplicadas:", dup_ids)

    # =====================================
    # 📥 EXPORTAR CON COLORES
    # =====================================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resultado.to_excel(writer, index=False, sheet_name="Resultado")

    output.seek(0)

    # 🔴 aplicar colores
    wb = load_workbook(output)
    ws = wb["Resultado"]

    rojo = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    amarillo = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

    for row in ws.iter_rows(min_row=2):
        irregular = str(row[resultado.columns.get_loc("Irregularidades")].value)

        if "CEDULA DUPLICADA" in irregular:
            for cell in row:
                cell.fill = rojo
        elif "SIN ID" in irregular:
            for cell in row:
                cell.fill = amarillo

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    # =====================================
    # ✅ RESULTADO
    # =====================================
    st.success("✅ Procesado correctamente")
    st.dataframe(resultado)

    st.download_button(
        "📥 Descargar Excel final",
        data=final_output.getvalue(),
        file_name="Resultado_Cursos_AVE.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
