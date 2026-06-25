import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Procesador AVE", layout="wide")

st.title("📊 Procesador de Cursos AVE")

# ======================================
# 🔘 SELECCIÓN DE TIPO
# ======================================
tipo_curso = st.radio(
    "Seleccione tipo de curso:",
    ["Curso CON nota", "Curso SIN nota"]
)

# ======================================
# 📂 CARGA DE ARCHIVOS
# ======================================
file_part = st.file_uploader("📂 Suba archivo PARTICIPANTES", type=["xlsx"])

if tipo_curso == "Curso CON nota":
    file_extra = st.file_uploader("📂 Suba archivo CALIFICACIONES", type=["xlsx"])
else:
    file_extra = st.file_uploader("📂 Suba archivo CERTIFICADOS", type=["xlsx"])

# ======================================
# 🚀 BOTÓN PROCESAR
# ======================================
if st.button("Procesar"):

    if file_part is not None and file_extra is not None:

        try:
            # ======================================
            # 📊 PARTICIPANTES
            # ======================================
            df_part = pd.read_excel(file_part)
            df_part.columns = df_part.columns.str.strip()

            st.subheader("📊 Participantes")
            st.dataframe(df_part.head())

            # ✅ CREAR CEDULA (CLAVE)
            df_part["Cedula"] = df_part["Número de ID"].astype(str).str.strip()

            # ✅ NOMBRE COMPLETO
            df_part["Nombre Completo"] = (
                df_part["Nombre"].astype(str).str.strip() + " " +
                df_part["Apellido(s)"].astype(str).str.strip()
            )

            # ✅ BASE FINAL
            df_final = df_part[[
                "Nombre Completo",
                "Cedula",
                "Departamento",
                "Institución"
            ]].copy()

            df_final.columns = [
                "Nombre Completo",
                "Cedula",
                "Cargo",
                "Seccional"
            ]

            # ======================================
            # 🟢 CASO: CURSO CON NOTA
            # ======================================
            if tipo_curso == "Curso CON nota":

                df_calif = pd.read_excel(file_extra)
                df_calif.columns = df_calif.columns.str.strip()

                st.subheader("📄 Calificaciones")
                st.dataframe(df_calif.head())

                # ✅ CREAR CEDULA
                df_calif["Cedula"] = df_calif["Número de ID"].astype(str).str.strip()

                df_calif_final = df_calif[[
                    "Cedula",
                    "Total del curso (Real)"
                ]].copy()

                df_calif_final.columns = ["Cedula", "Nota"]

                # ✅ CONVERTIR NOTA
                df_calif_final["Nota"] = pd.to_numeric(df_calif_final["Nota"], errors="coerce")

                # 🔄 CRUCE
                df_final = df_final.merge(df_calif_final, on="Cedula", how="left")

                # ✅ APROBACIÓN (tu lógica original)
                df_final["Aprobo"] = df_final["Nota"].apply(
                    lambda x: "No tiene nota" if pd.isna(x)
                    else "Sí" if x >= 3.5
                    else "No"
                )

                df_final["Estado"] = df_final["Nota"].apply(
                    lambda x: "No encontrado" if pd.isna(x) else "OK"
                )

            # ======================================
            # 🔵 CASO: CURSO SIN NOTA
            # ======================================
            else:

                df_cert = pd.read_excel(file_extra)
                df_cert.columns = df_cert.columns.str.strip()

                st.subheader("📄 Certificados")
                st.dataframe(df_cert.head())

                # ✅ CREAR CEDULA
                df_cert["Cedula"] = df_cert["Número de ID"].astype(str).str.strip()

                # ✅ MARCAR APROBADOS (C)
                df_cert["Aprobo_flag"] = 1

                # 🔄 CRUCE
                df_final = df_final.merge(
                    df_cert[["Cedula", "Aprobo_flag"]],
                    on="Cedula",
                    how="left"
                )

                # ✅ SIN NOTA
                df_final["Nota"] = ""

                # ✅ APROBACIÓN
                df_final["Aprobo"] = df_final["Aprobo_flag"].apply(
                    lambda x: "Sí" if pd.notnull(x) else "No"
                )

                df_final["Estado"] = df_final["Aprobo_flag"].apply(
                    lambda x: "OK" if pd.notnull(x) else "No encontrado"
                )

                df_final.drop(columns=["Aprobo_flag"], inplace=True)

            # ======================================
            # ✅ RESULTADO
            # ======================================
            st.success("✅ Procesado correctamente")
            st.dataframe(df_final)

            # ======================================
            # 📥 EXPORTAR
            # ======================================
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Resultado")

            st.download_button(
                label="📥 Descargar resultado",
                data=output.getvalue(),
                file_name="Resultado_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

    else:
        st.warning("⚠️ Debes subir los archivos requeridos")
