import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Cruce AVE - Participantes vs Calificaciones")

# ==============================
# 1. CARGA DE ARCHIVOS
# ==============================
file_part = st.file_uploader("Sube archivo PARTICIPANTES", type=["xlsx"])
file_calif = st.file_uploader("Sube archivo CALIFICACIONES", type=["xlsx"])

# ==============================
# 2. BOTÓN PROCESAR
# ==============================
if st.button("Procesar"):

    if file_part is not None and file_calif is not None:

        # Leer archivos
        df_part = pd.read_excel(file_part)
        df_calif = pd.read_excel(file_calif)

        # ==============================
        # PARTICIPANTES
        # ==============================
        df_part["Cedula"] = df_part["Número de ID"].astype(str).str.strip()

        df_part["Nombre Completo"] = (
            df_part["Nombre"].astype(str).str.strip() + " " +
            df_part["Apellido(s)"].astype(str).str.strip()
        )

        df_part_final = df_part[[
            "Nombre Completo",
            "Cedula",
            "Departamento",
            "Institución"
        ]].copy()

        df_part_final.columns = [
            "Nombre Completo",
            "Cedula",
            "Cargo",
            "Seccional"
        ]

        # ==============================
        # CALIFICACIONES
        # ==============================
        df_calif["Cedula"] = df_calif["Número de ID"].astype(str).str.strip()

        df_calif_final = df_calif[[
            "Cedula",
            "Total del curso (Real)"
        ]].copy()

        df_calif_final.columns = ["Cedula", "Nota"]

        # 🔥 convertir a número
        df_calif_final["Nota"] = pd.to_numeric(df_calif_final["Nota"], errors="coerce")

        # ==============================
        # CRUCE
        # ==============================
        df_final = df_part_final.merge(df_calif_final, on="Cedula", how="left")

        # ==============================
        # APROBACIÓN
        # ==============================
        df_final["Aprobo"] = df_final["Nota"].apply(
            lambda x: "No tiene nota" if pd.isna(x)
            else "Sí" if x >= 3.5
            else "No"
        )

        df_final["Estado"] = df_final["Nota"].apply(
            lambda x: "No encontrado" if pd.isna(x) else "OK"
        )

        # ==============================
        # MOSTRAR TABLA
        # ==============================
        st.success("✅ Procesado correctamente")
        st.dataframe(df_final)

        # ==============================
        # EXPORTAR A EXCEL
        # ==============================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Resultado")

        st.download_button(
            label="📥 Descargar resultado",
            data=output.getvalue(),
            file_name="Resultado_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("⚠️ Debes subir ambos archivos")
