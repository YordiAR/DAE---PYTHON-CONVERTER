import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Procesador AVE", layout="wide")

st.title("📊 Procesador de Cursos AVE")

# 🔘 Selector principal
tipo_curso = st.radio(
    "Seleccione tipo de curso:",
    ["Curso CON nota", "Curso SIN nota"]
)

# 📂 Carga archivos
archivo_part = st.file_uploader("📂 Suba archivo de Participantes", type=["xlsx"])

# Cambia según tipo
if tipo_curso == "Curso CON nota":
    archivo_extra = st.file_uploader("📂 Suba archivo de Calificaciones", type=["xlsx"])
else:
    archivo_extra = st.file_uploader("📂 Suba archivo de Certificados", type=["xlsx"])

# 🚀 PROCESAMIENTO
if archivo_part is not None and archivo_extra is not None:

    try:
        df_part = pd.read_excel(archivo_part)
        df_part.columns = df_part.columns.str.strip()

        st.subheader("📊 Participantes")
        st.dataframe(df_part.head())

        # 🔎 Detectar columnas
        col_nombre = next((c for c in df_part.columns if 'nombre' in c.lower()), None)
        col_cargo = next((c for c in df_part.columns if 'cargo' in c.lower()), None)
        col_seccional = next((c for c in df_part.columns if 'secc' in c.lower()), None)

        # 🧩 Base resultado
        resultado = pd.DataFrame()
        resultado['Nombre Completo'] = df_part[col_nombre] if col_nombre else ""
        resultado['Cedula'] = df_part['Cedula']
        resultado['Cargo'] = df_part[col_cargo] if col_cargo else ""
        resultado['Seccional'] = df_part[col_seccional] if col_seccional else ""

        # =====================================
        # 🟢 CURSO CON NOTA
        # =====================================
        if tipo_curso == "Curso CON nota":

            df_calif = pd.read_excel(archivo_extra)
            df_calif.columns = df_calif.columns.str.strip()

            st.subheader("📄 Calificaciones")
            st.dataframe(df_calif.head())

            col_id = next((c for c in df_calif.columns if 'ced' in c.lower()), None)
            col_nota = next((c for c in df_calif.columns if 'nota' in c.lower()), None)

            if not col_id or not col_nota:
                st.error("❌ El archivo de calificaciones no tiene estructura válida")
                st.stop()

            # 🔄 Cruce
            resultado = resultado.merge(
                df_calif[[col_id, col_nota]],
                left_on='Cedula',
                right_on=col_id,
                how='left'
            )

            resultado.rename(columns={col_nota: 'Nota'}, inplace=True)
            resultado.drop(columns=[col_id], inplace=True)

            # ✅ Reglas
            resultado['Aprobo'] = resultado['Nota'].apply(
                lambda x: 'Sí' if pd.notnull(x) and x >= 4 else 'No'
            )

            resultado['Estado'] = resultado['Nota'].apply(
                lambda x: 'OK' if pd.notnull(x) else 'No encontrado'
            )

        # =====================================
        # 🔵 CURSO SIN NOTA (CERTIFICADOS)
        # =====================================
        else:

            df_cert = pd.read_excel(archivo_extra)
            df_cert.columns = df_cert.columns.str.strip()

            st.subheader("📄 Certificados")
            st.dataframe(df_cert.head())

            # 🔎 Detectar columnas
            col_id = next((c for c in df_cert.columns if 'ced' in c.lower()), None)

            if not col_id:
                st.error("❌ El archivo de certificados debe contener cédula")
                st.stop()

            # ✅ Crear índice de certificados (aprobados)
            df_cert['Aprobo_flag'] = 1

            # 🔄 Cruce
            resultado = resultado.merge(
                df_cert[[col_id, 'Aprobo_flag']],
                left_on='Cedula',
                right_on=col_id,
                how='left'
            )

            resultado.drop(columns=[col_id], inplace=True)

            # ✅ Lógica SIN nota
            resultado['Nota'] = ""

            resultado['Aprobo'] = resultado['Aprobo_flag'].apply(
                lambda x: 'Sí' if pd.notnull(x) else 'No'
            )

            resultado['Estado'] = resultado['Aprobo_flag'].apply(
                lambda x: 'OK' if pd.notnull(x) else 'No encontrado'
            )

            resultado.drop(columns=['Aprobo_flag'], inplace=True)

        # 📊 Resultado final
        st.subheader("✅ Resultado final")
        st.dataframe(resultado)

        # 📥 Descargar
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            resultado.to_excel(writer, index=False, sheet_name='Resultado')

        output.seek(0)

        st.download_button(
            label="📥 Descargar resultado",
            data=output,
            file_name="Resultado_Curso.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("✅ Proceso completado correctamente")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
