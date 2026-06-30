import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =============================
# CONFIGURACIÓN FLEXIBLE
# =============================
participantes_file = "Participantes.xlsx"
calificaciones_file = "Calificaciones.xlsx"

# =============================
# CARGA DE DATOS
# =============================
df_part = pd.read_excel(participantes_file, engine="openpyxl")
df_cal = pd.read_excel(calificaciones_file, engine="openpyxl")

df_part.columns = df_part.columns.str.strip()
df_cal.columns = df_cal.columns.str.strip()

# =============================
# DETECTAR CÉDULA AUTOMÁTICAMENTE
# =============================
def find_cedula(df):
    for c in df.columns:
        if "cedula" in c.lower() or "documento" in c.lower():
            return c
    return None

cedula_part = find_cedula(df_part)
cedula_cal = find_cedula(df_cal)

if not cedula_part or not cedula_cal:
    raise ValueError("No se encontró columna de cédula en uno de los archivos")

df_part["_cedula"] = df_part[cedula_part].astype(str).str.strip()
df_cal["_cedula"] = df_cal[cedula_cal].astype(str).str.strip()

# =============================
# CÉDULAS VACÍAS Y DUPLICADAS
# =============================
mask_vacia = df_part["_cedula"].isin(["", "nan", "None"])
mask_dup = df_part["_cedula"].duplicated(keep=False) & ~mask_vacia

cedulas_vacias = df_part.loc[mask_vacia, "_cedula"].tolist()
cedulas_duplicadas = df_part.loc[mask_dup, "_cedula"].unique().tolist()

# =============================
# LIMPIEZA (SIN DUPLICADOS)
# =============================
df_clean = df_part[~mask_vacia].drop_duplicates(subset="_cedula", keep="first")

# =============================
# BUSCAR NOTA DE FORMA DINÁMICA
# =============================
score_col = None
for c in df_cal.columns:
    if c.strip() == "Total del curso (Real)":
        score_col = c
        break

if not score_col:
    raise ValueError("No se encontró 'Total del curso (Real)'")

# =============================
# CRUCE
# =============================
df_merge = df_clean.merge(
    df_cal[["_cedula", score_col]],
    on="_cedula",
    how="left"
)

# =============================
# ESTADO DEL PARTICIPANTE
# =============================
df_merge["estado"] = df_merge[score_col].apply(
    lambda x: "CERTIFICADO" if pd.notna(x) and x >= 70
    else ("NO CERTIFICADO" if pd.notna(x) else "SIN NOTA")
)

# =============================
# MÉTRICAS
# =============================
resumen = {
    "total_participantes": len(df_part),
    "sin_duplicados": len(df_clean),
    "cedulas_duplicadas": len(cedulas_duplicadas),
    "cedulas_vacias": len(cedulas_vacias),
    "certificados": int((df_merge["estado"] == "CERTIFICADO").sum()),
    "no_certificados": int((df_merge["estado"] == "NO CERTIFICADO").sum())
}

print(resumen)

# =============================
# EXPORTAR PARA STREAMLIT
# =============================
df_merge.to_excel("resultado_cruzado.xlsx", index=False, engine="openpyxl")

pd.DataFrame(df_merge).to_json("data_streamlit.json", orient="records", force_ascii=False)

pd.DataFrame([{
    **resumen,
    "lista_duplicadas": cedulas_duplicadas,
    "lista_vacias": cedulas_vacias
}]).to_json("resumen.json", orient="records", force_ascii=False)

# =============================
# MARCAR EN EXCEL (ROJO CLARO)
# =============================
wb = load_workbook("resultado_cruzado.xlsx")
ws = wb.active

fill_red = PatternFill(start_color="FFF4CCCC", end_color="FFF4CCCC", fill_type="solid")

col_idx = None
for i, cell in enumerate(ws[1], start=1):
    if cell.value == "_cedula":
        col_idx = i
        break

for row in range(2, ws.max_row + 1):
    ced = str(ws.cell(row=row, column=col_idx).value).strip()

    if ced in cedulas_duplicadas or ced in ["", "nan", "None"]:
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill_red

wb.save("resultado_validado.xlsx")
