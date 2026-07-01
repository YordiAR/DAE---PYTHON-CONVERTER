import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

st.set_page_config(page_title="Procesador AVE - Blindado", layout="wide")
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
# 🔎 DETECTAR COLUMNA DE ID (MEJORADO)
# =====================================
def find_id_column(df):
    posibles = [
        "cedula", "cédula",
        "numero de id", "número de id",
        "no identificacion", "n° identificacion",
        "identificacion", "identificación",
        "documento", "documento de identidad",
        "id"
    ]
    for k in posibles:
        col = next((c for c in df.columns if k in c.lower()), None)
        if col:
            return col
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
# 🔘 TIPO DE CURSO
# =====================================
tipo_curso = st.radio(
    "Seleccione tipo de curso:",
    ["Curso CON nota", "Curso SIN nota"]
)

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

    try:
        # =====================================
        # 📊 PARTICIPANTES
        # =====================================
        df_part = pd.read_excel(file_part)
        df_part.columns = df_part.columns.str.strip()

        col_id_part = find_id_column(df_part)
        col_nombre_part = find_column(df_part, ["nombre"])
        col_apellido_part = find_column(df_part, ["apellido"])

        if not col_id_part:
            st.error("❌ No se encontró columna de identificación válida")
            st.stop()

        # =====================================
        # 🔥 LIMPIEZA REAL DE CÉDULA
        # =====================================
        df_part["Cedula"] = df_part[col_id_part].astype(str)

        # quitar todo lo que no sea número
        df_part["Cedula"] = df_part["Cedula"].str.replace(r"\D", "", regex=True)

        # eliminar filas sin ID válido
        df_part = df_part[df_part["Cedula"].str.len() >= 6].copy()

        # =====================================
        # 👤 NOMBRES
        # =====================================
        if col_nombre_part and col_apellido_part:
            df_part["Nombres y Apellidos"] = (
                normalize_text(df_part[col_nombre_part]) + " " +
                normalize_text(df_part[col_apellido_part])
            )
        elif col_nombre_part:
            df_part["Nombres y Apellidos"] = normalize_text(df_part[col_nombre_part])
        else:
            df_part["Nombres y Apellidos"] = ""

        # =====================================
        # ⚠️ DUPLICADOS REALES (POR CÉDULA LIMPIA)
        # =====================================
        df_part["SIN_ID"] = df_part["Cedula"].isin(["", "NAN", "NONE"])
        df_part["DUPLICADO"] = df_part.duplicated(subset=["Cedula"], keep=False)

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

            df_calif["Cedula"] = df_calif[col_id_calif].astype(str)
            df_calif["Cedula"] = df_calif["Cedula"].str.replace(r"\D", "", regex=True)
            df_calif["Nota"] = pd.to_numeric(df_calif[col_nota], errors="coerce")

            # 🔴 quitar duplicados (mejor nota)
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
        # 🔵 SIN NOTA
        # =====================================
        else:

            df_aprob = pd.read_excel(file_extra)
            df_aprob.columns = df_aprob.columns.str.strip()

            col_id_aprob = find_id_column(df_aprob)

            if not col_id_aprob:
                st.error("❌ Aprobados: no se encontró columna de ID válida")
                st.stop()

            df_aprob["Cedula"] = df_aprob[col_id_aprob].astype(str)
            df_aprob["Cedula"] = df_aprob["Cedula"].str.replace(r"\D", "", regex=True)

            df_aprob = df_aprob.drop_duplicates("Cedula")

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

        resultado.loc[df_part["SIN_ID"].values, "Irregularidades"] += "SIN ID; "
        resultado.loc[df_part["DUPLICADO"].values, "Irregularidades"] += "CEDULA DUPLICADA; "
        resultado.loc[resultado["Aprobo"] == "No", "Irregularidades"] += "NO ENCONTRADO; "

        # =====================================
        # 📊 CONTROL
        # =====================================
        st.subheader("📊 Control de calidad")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total", len(df_part))
        col2.metric("Sin ID", int(df_part["SIN_ID"].sum()))
        col3.metric("Duplicados reales", int(df_part["DUPLICADO"].sum()))
        col4.metric("No encontrados", int((resultado["Aprobo"] == "No").sum()))

        st.write("📌 Columna ID detectada:", col_id_part)
        st.write("🔎 Ejemplo de cédulas:", df_part["Cedula"].head(10).tolist())

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
            irregular = str(row[col_irreg].value)

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

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
