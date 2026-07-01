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

    try:
        # =====================================
        # 📊 PARTICIPANTES
        # =====================================
        df_part = pd.read_excel(file_part)
        df_part.columns = df_part.columns.str.strip()

        # 🔎 detectar columnas reales (FIX CLAVE)
        col_id_part = next((c for c in df_part.columns if "id" in c.lower()), None)
        col_nombre_part = next((c for c in df_part.columns if "nombre" in c.lower()), None)
        col_apellido_part = next((c for c in df_part.columns if "apellido" in c.lower()), None)

        if not col_id_part:
            st.error("❌ El archivo de participantes no tiene columna Número de ID")
            st.stop()

        # ✅ construir campos
        df_part["Cedula"] = df_part[col_id_part].astype(str).str.strip()

        if col_nombre_part and col_apellido_part:
            df_part["Nombres y Apellidos"] = (
                df_part[col_nombre_part].astype(str).str.strip() + " " +
                df_part[col_apellido_part].astype(str).str.strip()
            )
        elif col_nombre_part:
            df_part["Nombres y Apellidos"] = df_part[col_nombre_part].astype(str).str.strip()
        else:
            df_part["Nombres y Apellidos"] = ""

        # =====================================
        # ⚠️ VALIDACIONES PARTICIPANTES
        # =====================================
        df_part["SIN_ID"] = df_part["Cedula"].isin(["", "nan", None])
        df_part["DUPLICADO"] = df_part.duplicated(subset=["Cedula"], keep=False)

        # =====================================
        # 🧠 BASE FINAL
        # =====================================
        base = df_part[[
            "Nombres y Apellidos",
            "Cedula",
            "Departamento",
            "Institución"
        ]].copy()

        base.columns = ["Nombres y Apellidos", "Cedula", "Cargo", "Dependencia"]

        # =====================================
        # 🟢 CURSO CON NOTA
        # =====================================
        if tipo_curso == "Curso CON nota":

            df_calif = pd.read_excel(file_extra)
            df_calif.columns = df_calif.columns.str.strip()

            # 🔎 FIX: detectar nombres de columnas
            col_id_calif = next((c for c in df_calif.columns if "id" in c.lower()), None)
            col_nombre_calif = next((c for c in df_calif.columns if "nombre" in c.lower()), None)
            col_apellido_calif = next((c for c in df_calif.columns if "apellido" in c.lower()), None)
            col_nota = next((c for c in df_calif.columns if "total del curso" in c.lower()), None)

            if not col_id_calif or not col_nota:
                st.error("❌ Estructura inválida en calificaciones")
                st.stop()

            df_calif["Cedula"] = df_calif[col_id_calif].astype(str).str.strip()

            if col_nombre_calif and col_apellido_calif:
                df_calif["Nombres y Apellidos"] = (
                    df_calif[col_nombre_calif].astype(str).str.strip() + " " +
                    df_calif[col_apellido_calif].astype(str).str.strip()
                )
            else:
                df_calif["Nombres y Apellidos"] = ""

            df_calif["Nota"] = pd.to_numeric(df_calif[col_nota], errors="coerce")

            # 🔄 CRUCE (ID + NOMBRES)
            resultado = base.merge(
                df_calif[["Cedula", "Nombres y Apellidos", "Nota"]],
                on=["Cedula", "Nombres y Apellidos"],
                how="left"
            )

            resultado["Aprobo"] = resultado["Nota"].apply(
                lambda x: "Sí" if pd.notnull(x) and x >= 3.5 else "No"
            )

        # =====================================
        # 🔵 CURSO SIN NOTA
        # =====================================
        else:

            df_aprob = pd.read_excel(file_extra)
            df_aprob.columns = df_aprob.columns.str.strip()

            # 🔎 FIX CLAVE (no asumir "Nombre")
            col_id_aprob = next((c for c in df_aprob.columns if "id" in c.lower()), None)
            col_nombre_aprob = next((c for c in df_aprob.columns if "nombre" in c.lower()), None)
            col_apellido_aprob = next((c for c in df_aprob.columns if "apellido" in c.lower()), None)

            if not col_id_aprob:
                st.error("❌ El archivo de aprobados no tiene Número de ID")
                st.stop()

            df_aprob["Cedula"] = df_aprob[col_id_aprob].astype(str).str.strip()

            if col_nombre_aprob and col_apellido_aprob:
                df_aprob["Nombres y Apellidos"] = (
                    df_aprob[col_nombre_aprob].astype(str).str.strip() + " " +
                    df_aprob[col_apellido_aprob].astype(str).str.strip()
                )
            elif col_nombre_aprob:
                df_aprob["Nombres y Apellidos"] = df_aprob[col_nombre_aprob].astype(str).str.strip()
            else:
                df_aprob["Nombres y Apellidos"] = ""

            df_aprob["Aprobo_flag"] = 1

            # 🔄 CRUCE (solo ID + Nombres)
            resultado = base.merge(
                df_aprob[["Cedula", "Nombres y Apellidos", "Aprobo_flag"]],
                on=["Cedula", "Nombres y Apellidos"],
                how="left"
            )

            resultado["Nota"] = ""

            resultado["Aprobo"] = resultado["Aprobo_flag"].apply(
                lambda x: "Sí" if pd.notnull(x) else "No"
            )

            resultado.drop(columns=["Aprobo_flag"], inplace=True)

        # =====================================
        # ⚠️ IRREGULARIDADES
        # =====================================
        resultado["Irregularidades"] = ""

        resultado.loc[df_part["SIN_ID"], "Irregularidades"] += "SIN ID; "
        resultado.loc[df_part["DUPLICADO"], "Irregularidades"] += "CEDULA DUPLICADA; "
        resultado.loc[resultado["Aprobo"] == "No", "Irregularidades"] += "NO ENCONTRADO; "

        # =====================================
        # 📊 RESUMEN
        # =====================================
        st.subheader("📊 Resumen")
        st.write("Total participantes:", len(df_part))
        st.write("Sin ID:", df_part["SIN_ID"].sum())
        st.write("Duplicados:", df_part["DUPLICADO"].sum())
        st.write("Aprobados:", (resultado["Aprobo"] == "Sí").sum())

        st.write("Cédulas sin ID:", df_part[df_part["SIN_ID"]]["Cedula"].tolist())
        st.write("Cédulas duplicadas:", df_part[df_part["DUPLICADO"]]["Cedula"].tolist())

        # =====================================
        # 📥 EXPORTAR CON COLORES
        # =====================================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            resultado.to_excel(writer, index=False, sheet_name="Resultado")

        output.seek(0)

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

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
