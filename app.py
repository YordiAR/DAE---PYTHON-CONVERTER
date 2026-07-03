 

# ========================================================== 

# APP CURSOS - STREAMLIT 

# Autor: ChatGPT 

# ========================================================== 

 

import streamlit as st 

import pandas as pd 

import numpy as np 

 

from io import BytesIO 

 

from openpyxl import Workbook 

from openpyxl.styles import PatternFill 

from openpyxl.styles import Font 

from openpyxl.styles import Alignment 

from openpyxl.styles import Border 

from openpyxl.styles import Side 

from openpyxl.utils.dataframe import dataframe_to_rows 

 

# ---------------------------------------------------------- 

# CONFIGURACIÓN STREAMLIT 

# ---------------------------------------------------------- 

 

st.set_page_config( 

    page_title="Procesador de Cursos", 

    page_icon="📚", 

    layout="wide" 

) 

 

st.title("📚 Procesador de Cursos") 

st.caption("Cursos con nota y sin nota") 

 

# ---------------------------------------------------------- 

# COLORES 

# ---------------------------------------------------------- 

 

COLOR_ROJO = "FFC7CE" 

COLOR_AMARILLO = "FFF2CC" 

COLOR_VERDE = "C6EFCE" 

COLOR_GRIS = "D9D9D9" 

 

# ---------------------------------------------------------- 

# COLUMNAS ESPERADAS 

# ---------------------------------------------------------- 

 

COLUMNAS_PARTICIPANTES = [ 

    "Número de ID", 

    "Nombre", 

    "Apellido(s)", 

    "Departamento", 

    "Institución" 

] 

 

COLUMNAS_APROBADOS = [ 

    "Número de ID" 

] 

 

# ---------------------------------------------------------- 

# FUNCIONES AUXILIARES 

# ---------------------------------------------------------- 

 

def limpiar_texto(valor): 

    """ 

    Limpia texto para facilitar comparaciones. 

    """ 

 

    if pd.isna(valor): 

        return "" 

 

    return ( 

        str(valor) 

        .strip() 

        .upper() 

        .replace("Á", "A") 

        .replace("É", "E") 

        .replace("Í", "I") 

        .replace("Ó", "O") 

        .replace("Ú", "U") 

    ) 

 

 

def nombre_completo(df): 

 

    nombre = df["Nombre"].fillna("").astype(str) 

 

    apellido = df["Apellido(s)"].fillna("").astype(str) 

 

    return ( 

        nombre.str.strip() 

        + " " 

        + apellido.str.strip() 

    ).str.upper().str.strip() 

 

 

# ---------------------------------------------------------- 

# DETECTAR COLUMNA DE NOTA 

# ---------------------------------------------------------- 

 

def obtener_columna_nota(df): 

 

    posibles = [ 

 

        "Total del curso (Real)", 

 

        "Total del Curso (Real)", 

 

        "Nota", 

 

        "Nota Final", 

 

        "Calificación", 

 

        "Calificacion" 

 

    ] 

 

    for c in df.columns: 

 

        if c in posibles: 

            return c 

 

    for c in df.columns: 

 

        texto = c.lower() 

 

        if "total del curso" in texto: 

            return c 

 

        if "nota" in texto: 

            return c 

 

        if "calificacion" in texto: 

            return c 

 

    return None 

 

# ---------------------------------------------------------- 

# LECTURA DE EXCEL 

# ---------------------------------------------------------- 

 

def cargar_excel(archivo): 

 

    try: 

 

        df = pd.read_excel(archivo) 

 

        df.columns = ( 

            df.columns 

            .astype(str) 

            .str.strip() 

            .str.replace("\n", " ") 

        ) 

 

        return df 

 

    except Exception as e: 

 

        st.error(f"Error leyendo archivo: {e}") 

 

        return None 

 

 

# ---------------------------------------------------------- 

# VALIDAR COLUMNAS 

# ---------------------------------------------------------- 

 

def validar_participantes(df): 

 

    faltantes = [] 

 

    for col in COLUMNAS_PARTICIPANTES: 

 

        if col not in df.columns: 

            faltantes.append(col) 

 

    return faltantes 

 

 

def validar_aprobados(df): 

 

    faltantes = [] 

 

    for col in COLUMNAS_APROBADOS: 

 

        if col not in df.columns: 

            faltantes.append(col) 

 

    return faltantes 

 

 

def validar_calificaciones(df): 

 

    nota = obtener_columna_nota(df) 

 

    if nota is None: 

 

        return False 

 

    return True 

 

 

# ---------------------------------------------------------- 

# NORMALIZAR PARTICIPANTES 

# ---------------------------------------------------------- 

 

def preparar_participantes(df): 

 

    df = df.copy() 

 

    df["Número de ID"] = ( 

        df["Número de ID"] 

        .astype(str) 

        .replace("nan", "") 

        .str.strip() 

    ) 

 

    df["Nombre Completo"] = nombre_completo(df) 

 

    df["Observaciones"] = "" 

 

    return df 

 

 

# ---------------------------------------------------------- 

# NORMALIZAR CALIFICACIONES 

# ---------------------------------------------------------- 

 

def preparar_calificaciones(df): 

 

    df = df.copy() 

 

    nota = obtener_columna_nota(df) 

 

    df["Nota"] = pd.to_numeric( 

        df[nota], 

        errors="coerce" 

    ) 

 

    if "Número de ID" in df.columns: 

 

        df["Número de ID"] = ( 

            df["Número de ID"] 

            .astype(str) 

            .replace("nan", "") 

            .str.strip() 

        ) 

 

    df["Nombre Completo"] = nombre_completo(df) 

 

    return df 

 

 

# ---------------------------------------------------------- 

# NORMALIZAR APROBADOS 

# ---------------------------------------------------------- 

 

def preparar_aprobados(df): 

 

    df = df.copy() 

 

    df["Número de ID"] = ( 

        df["Número de ID"] 

        .astype(str) 

        .replace("nan", "") 

        .str.strip() 

    ) 

 

    return df 

# ========================================================== 

# VALIDACIONES 

# ========================================================== 

 

def validar_ids_vacios(df): 

 

    df = df.copy() 

 

    mascara = ( 

        df["Número de ID"].isna() 

        | (df["Número de ID"] == "") 

        | (df["Número de ID"].astype(str).str.strip() == "") 

    ) 

 

    df.loc[mascara, "Observaciones"] = ( 

        df.loc[mascara, "Observaciones"] 

        + "ID vacío; " 

    ) 

 

    return df, mascara 

 

 

def validar_duplicados(df): 

 

    df = df.copy() 

 

    mascara = df["Número de ID"].duplicated(keep=False) 

 

    mascara &= df["Número de ID"] != "" 

 

    df.loc[mascara, "Observaciones"] = ( 

        df.loc[mascara, "Observaciones"] 

        + "Cédula duplicada; " 

    ) 

 

    return df, mascara 

 

 

# ========================================================== 

# CRUCE POR ID 

# ========================================================== 

 

def cruzar_por_id(participantes, resultados): 

 

    columnas = [ 

        "Número de ID", 

        "Nota", 

        "Nombre Completo" 

    ] 

 

    disponibles = [ 

        c for c in columnas if c in resultados.columns 

    ] 

 

    merge = participantes.merge( 

        resultados[disponibles], 

        on="Número de ID", 

        how="left", 

        suffixes=("", "_resultado") 

    ) 

 

    return merge 

 

 

# ========================================================== 

# CRUCE POR NOMBRE 

# ========================================================== 

 

def completar_por_nombre(df, resultados): 

 

    if "Nombre Completo" not in resultados.columns: 

        return df 

 

    resultados_nombre = resultados[ 

        ["Nombre Completo", "Nota"] 

    ].drop_duplicates() 

 

    sin_nota = df["Nota"].isna() 

 

    temporal = df.loc[sin_nota].merge( 

        resultados_nombre, 

        on="Nombre Completo", 

        how="left", 

        suffixes=("", "_nombre") 

    ) 

 

    if "Nota_nombre" in temporal.columns: 

 

        df.loc[sin_nota, "Nota"] = temporal["Nota_nombre"].values 

 

    return df 

 

 

# ========================================================== 

# APROBACIÓN CURSOS CON NOTA 

# ========================================================== 

 

def calcular_aprobados_con_nota(df): 

 

    df = df.copy() 

 

    df["Aprobó"] = np.where( 

        df["Nota"] >= 3.5, 

        "Sí", 

        "No" 

    ) 

 

    return df 

 

 

# ========================================================== 

# APROBACIÓN CURSOS SIN NOTA 

# ========================================================== 

 

def calcular_aprobados_sin_nota(participantes, aprobados): 

 

    participantes = participantes.copy() 

 

    participantes["Aprobó"] = np.where( 

 

        participantes["Número de ID"].isin( 

            aprobados["Número de ID"] 

        ), 

 

        "Sí", 

 

        "No" 

 

    ) 

 

    participantes["Nota"] = "" 

 

    return participantes 

 

 

# ========================================================== 

# OBSERVACIONES 

# ========================================================== 

 

def limpiar_observaciones(df): 

 

    df["Observaciones"] = ( 

        df["Observaciones"] 

        .str.strip() 

        .str.rstrip(";") 

    ) 

 

    return df 

 

 

# ========================================================== 

# DATAFRAME FINAL 

# ========================================================== 

 

def construir_dataframe(df): 

 

    df = df.copy() 

 

    df.rename( 

        columns={ 

            "Número de ID": "Cédula", 

            "Departamento": "Cargo", 

            "Institución": "Dependencia", 

            "Nombre Completo": "Nombres y apellidos" 

        }, 

        inplace=True 

    ) 

 

    columnas = [ 

 

        "Nombres y apellidos", 

 

        "Cédula", 

 

        "Cargo", 

 

        "Dependencia", 

 

        "Nota", 

 

        "Aprobó", 

 

        "Observaciones" 

 

    ] 

 

    return df[columnas] 

 

 

# ========================================================== 

# RESUMEN 

# ========================================================== 

 

def generar_resumen(df): 

 

    ids_vacios = ( 

        df["Observaciones"] 

        .str.contains("ID vacío",
                      case=False,
                      na=False)

    ).sum() 

 

    duplicados = ( 

        df["Cédula"] 

        .duplicated(keep=False) 

    ) 

 

    total_duplicados = duplicados.sum() 

 

    aprobados = ( 

        df["Aprobó"] == "Sí" 

    ).sum() 

    aprobados_duplicados = ( 

        (df["Aprobó"] == "Sí") & duplicados 

    ).sum()

 

    resumen = { 

 

        "Participantes": len(df), 

 

        "Aprobados": aprobados, 

 

        "Reprobados": len(df) - aprobados, 

 

        "IDs Vacíos": ids_vacios, 

 

        "Duplicados": total_duplicados,



        "Aprobados Duplicados": aprobados_duplicados

 

    } 

 

    return resumen 

 

 

# ========================================================== 

# INCONSISTENCIAS 

# ========================================================== 

 

def obtener_inconsistencias(df): 

 

    inconsistencias = df[ 

 

        (df["Observaciones"] != "") 

 

    ].copy() 

 

    return inconsistencias 

# ========================================================== 

# EXPORTAR A EXCEL 

# ========================================================== 

 

def exportar_excel(df_final, resumen, inconsistencias): 

 

    wb = Workbook() 

 

    ws = wb.active 

    ws.title = "Reporte_Final" 

 

    # ------------------------- 

    # Estilos 

    # ------------------------- 

 

    encabezado_fill = PatternFill( 

        fill_type="solid", 

        fgColor=COLOR_GRIS 

    ) 

 

    rojo_fill = PatternFill( 

        fill_type="solid", 

        fgColor=COLOR_ROJO 

    ) 

 

    amarillo_fill = PatternFill( 

        fill_type="solid", 

        fgColor=COLOR_AMARILLO 

    ) 

 

    verde_fill = PatternFill( 

        fill_type="solid", 

        fgColor=COLOR_VERDE 

    ) 

 

    borde = Border( 

        left=Side(style="thin"), 

        right=Side(style="thin"), 

        top=Side(style="thin"), 

        bottom=Side(style="thin") 

    ) 

 

    # ------------------------- 

    # Reporte principal 

    # ------------------------- 

 

    for fila in dataframe_to_rows(df_final, 

                                  index=False, 

                                  header=True): 

 

        ws.append(fila) 

 

    # Encabezado 

    for celda in ws[1]: 

 

        celda.fill = encabezado_fill 

        celda.font = Font(bold=True) 

        celda.alignment = Alignment(horizontal="center") 

        celda.border = borde 

 

    # Filas 

    for fila in ws.iter_rows(min_row=2): 

 

        observacion = str(fila[6].value) 

 

        if "ID vacío" in observacion: 

 

            for c in fila: 

                c.fill = amarillo_fill 

 

        if "duplicada" in observacion.lower(): 

 

            for c in fila: 

                c.fill = rojo_fill 

 

        if fila[5].value == "Sí": 

 

            fila[5].fill = verde_fill 

 

        for c in fila: 

            c.border = borde 

 

    # ------------------------- 

    # Auto ancho columnas 

    # ------------------------- 

 

    for columna in ws.columns: 

 

        largo = 0 

 

        letra = columna[0].column_letter 

 

        for celda in columna: 

 

            try: 

 

                if len(str(celda.value)) > largo: 

 

                    largo = len(str(celda.value)) 

 

            except: 

 

                pass 

 

        ws.column_dimensions[letra].width = largo + 4 

 

    ws.freeze_panes = "A2" 

 

    ws.auto_filter.ref = ws.dimensions 

 

    # ===================================================== 

    # HOJA RESUMEN 

    # ===================================================== 

 

    resumen_ws = wb.create_sheet("Resumen") 

 

    resumen_ws.append(["Indicador", "Valor"]) 

 

    for c in resumen_ws[1]: 

 

        c.fill = encabezado_fill 

        c.font = Font(bold=True) 

        c.border = borde 

 

    for k, v in resumen.items(): 

 

        resumen_ws.append([k, v]) 

 

    for fila in resumen_ws.iter_rows(): 

 

        for c in fila: 

            c.border = borde 

 

    resumen_ws.column_dimensions["A"].width = 35 

    resumen_ws.column_dimensions["B"].width = 15 

 

    # ===================================================== 

    # HOJA INCONSISTENCIAS 

    # ===================================================== 

 

    inc_ws = wb.create_sheet("Inconsistencias") 

 

    for fila in dataframe_to_rows( 

            inconsistencias, 

            index=False, 

            header=True): 

 

        inc_ws.append(fila) 

 

    for celda in inc_ws[1]: 

 

        celda.fill = encabezado_fill 

        celda.font = Font(bold=True) 

        celda.border = borde 

 

    for fila in inc_ws.iter_rows(min_row=2): 

 

        obs = str(fila[6].value) 

 

        if "ID vacío" in obs: 

 

            for c in fila: 

                c.fill = amarillo_fill 

 

        if "duplicada" in obs.lower(): 

 

            for c in fila: 

                c.fill = rojo_fill 

 

        for c in fila: 

            c.border = borde 

 

    for columna in inc_ws.columns: 

 

        largo = 0 

 

        letra = columna[0].column_letter 

 

        for celda in columna: 

 

            try: 

 

                if len(str(celda.value)) > largo: 

                    largo = len(str(celda.value)) 

 

            except: 

                pass 

 

        inc_ws.column_dimensions[letra].width = largo + 4 

 

    # ===================================================== 

    # Guardar en memoria 

    # ===================================================== 

 

    salida = BytesIO() 

 

    wb.save(salida) 

 

    salida.seek(0) 

 

    return salida 

 

# ========================================================== 

# INTERFAZ STREAMLIT 

# ========================================================== 

 

st.divider() 

 

tipo_curso = st.radio( 

    "Seleccione el tipo de curso", 

    ["Con nota", "Sin nota"], 

    horizontal=True 

) 

 

st.divider() 

 

participantes_file = st.file_uploader( 

    "📄 Archivo de Participantes", 

    type=["xlsx"] 

) 

 

if tipo_curso == "Con nota": 

 

    resultados_file = st.file_uploader( 

        "📄 Archivo de Calificaciones", 

        type=["xlsx"] 

    ) 

 

else: 

 

    resultados_file = st.file_uploader( 

        "📄 Archivo de Aprobados", 

        type=["xlsx"] 

    ) 

 

procesar = st.button( 

    "🚀 Procesar", 

    use_container_width=True 

) 

 

# ========================================================== 

# PROCESAMIENTO 

# ========================================================== 

 

if procesar: 

 

    if participantes_file is None: 

 

        st.error("Debe cargar el archivo de participantes.") 

        st.stop() 

 

    if resultados_file is None: 

 

        st.error("Debe cargar el segundo archivo.") 

        st.stop() 

 

    barra = st.progress(0) 

 

    try: 

 

        # ----------------------------------------- 

        # Lectura 

        # ----------------------------------------- 

 

        barra.progress(10) 

 

        participantes = cargar_excel(participantes_file) 

 

        resultados = cargar_excel(resultados_file) 

 

        # ----------------------------------------- 

        # Validaciones 

        # ----------------------------------------- 

 

        barra.progress(20) 

 

        faltantes = validar_participantes(participantes) 

 

        if len(faltantes) > 0: 

 

            st.error( 

                f"Faltan columnas en Participantes:\n{faltantes}" 

            ) 

 

            st.stop() 

 

        if tipo_curso == "Con nota": 

 

            if not validar_calificaciones(resultados): 

 

                st.error( 

                    "No fue posible encontrar la columna de nota." 

                ) 

 

                st.stop() 

 

        else: 

 

            faltantes = validar_aprobados(resultados) 

 

            if len(faltantes) > 0: 

 

                st.error( 

                    f"Faltan columnas en archivo de aprobados:\n{faltantes}" 

                ) 

 

                st.stop() 

 

        # ----------------------------------------- 

        # Preparación 

        # ----------------------------------------- 

 

        barra.progress(35) 

 

        participantes = preparar_participantes(participantes) 

 

        if tipo_curso == "Con nota": 

 

            resultados = preparar_calificaciones(resultados) 

 

        else: 

 

            resultados = preparar_aprobados(resultados) 

 

        # ----------------------------------------- 

        # Validaciones de participantes 

        # ----------------------------------------- 

 

        barra.progress(45) 

 

        participantes, ids_vacios = validar_ids_vacios( 

            participantes 

        ) 

 

        participantes, duplicados = validar_duplicados( 

            participantes 

        ) 

 

        # ----------------------------------------- 

        # Procesamiento 

        # ----------------------------------------- 

 

        barra.progress(65) 

 

        if tipo_curso == "Con nota": 

 

            df = cruzar_por_id( 

                participantes, 

                resultados 

            ) 

 

            df = completar_por_nombre( 

                df, 

                resultados 

            ) 

 

            df = calcular_aprobados_con_nota( 

                df 

            ) 

 

        else: 

 

            df = calcular_aprobados_sin_nota( 

                participantes, 

                resultados 

            ) 

 

        # ----------------------------------------- 

        # Observaciones 

        # ----------------------------------------- 

 

        barra.progress(75) 

 

        df = limpiar_observaciones(df) 

 

        # ----------------------------------------- 

        # DataFrame Final 

        # ----------------------------------------- 

 

        barra.progress(85) 

 

        reporte = construir_dataframe(df) 

 

        resumen = generar_resumen(reporte) 

 

        inconsistencias = obtener_inconsistencias( 

            reporte 

        ) 

 

        # ----------------------------------------- 

        # Excel 

        # ----------------------------------------- 

 

        barra.progress(95) 

 

        archivo = exportar_excel( 

            reporte, 

            resumen, 

            inconsistencias 

        ) 

 

        barra.progress(100) 

 

        st.success("Proceso finalizado correctamente.") 

 

        # ===================================================== 

        # INDICADORES 

        # ===================================================== 

 

        st.divider() 

 

        c1, c2, c3, c4, c5 = st.columns(5) 

 

        c1.metric( 

            "Participantes", 

            resumen["Participantes"] 

        ) 

 

        c2.metric( 

            "Aprobados", 

            resumen["Aprobados"] 

        ) 

 

        c3.metric( 

            "IDs Vacíos", 

            resumen["IDs Vacíos"] 

        ) 

 

        c4.metric( 

            "Duplicados", 

            resumen["Duplicados"] 

        ) 


        c5.metric( 

            "Aprobados Duplicados", 

            resumen["Aprobados Duplicados"] 

        )

 

        st.divider() 

 

        # ===================================================== 

        # CÉDULAS DUPLICADAS 

        # ===================================================== 

 

        duplicadas = reporte[ 

            reporte["Cédula"].duplicated(keep=False) 

        ]["Cédula"].unique() 

 

        vacias = reporte[ 

            reporte["Observaciones"].str.contains("ID vacío",
                      case=False,
                      na=False)

        ] 

 

        col1, col2 = st.columns(2) 

 

        with col1: 

 

            st.subheader("Cédulas duplicadas") 

 

            if len(duplicadas): 

 

                st.write(list(duplicadas)) 

 

            else: 

 

                st.success("No existen.") 

 

        with col2: 

 

            st.subheader("Registros sin ID") 

 

            if len(vacias): 

 

                st.dataframe(vacias) 

 

            else: 

 

                st.success("No existen.") 

 

        st.divider() 

 

        # ===================================================== 

        # PREVISUALIZACIÓN 

        # ===================================================== 

 

        st.subheader("Vista previa") 

 

        st.dataframe( 

            reporte, 

            use_container_width=True, 

            hide_index=True 

        ) 

 

        st.download_button( 

 

            label="📥 Descargar Excel", 

 

            data=archivo, 

 

            file_name="Reporte_Final.xlsx", 

 

            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 

 

            use_container_width=True 

 

        ) 

 

    except Exception as e: 

 

        st.exception(e) 

 
