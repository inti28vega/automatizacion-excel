import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 1. SIMULAR DATOS CON ERRORES Y ANOMALÍAS
datos_sucios = pd.DataFrame({
    "ID_Transaccion": [101, 102, 103, 104, 105, 105],  # 105 está duplicado
    "Vendedor": ["Carlos", "Ana", "  Pedro ", None, "Carlos", "Carlos"], # Pedro con espacios, None es nulo
    "Monto": [1500.0, -200.0, 800.0, 1200.0, 0.0, 1500.0], # Montos negativos y ceros
    "Estado": ["Completado", "Completado", "Completado", "Completado", "Pendiente", "Completado"]
})

datos_sucios.to_csv("ventas_raw_sucias.csv", index=False)

# 2. PROCESO DE AUDITORÍA Y VALIDACIÓN DE DATOS
df = pd.read_csv("ventas_raw_sucias.csv")

# A. Limpieza de texto (Quitar espacios en blanco al inicio/final)
df["Vendedor"] = df["Vendedor"].astype(str).str.strip()

# B. Identificar filas con errores/anomalías
errores_nulos = df[df["Vendedor"] == "None"]
errores_montos = df[df["Monto"] <= 0]
errores_duplicados = df[df.duplicated(subset=["ID_Transaccion"], keep=False)]

# Guardar un Log de Auditoría en CSV con los registros inconsistentes
df_errores = pd.concat([errores_nulos, errores_montos, errores_duplicados]).drop_duplicates()
df_errores.to_csv("auditoria_errores_detectados.csv", index=False)

# C. FILTRAR DATA LIMPIA (Descartar anomalías)
df_limpio = df[
    (df["Vendedor"] != "None") & 
    (df["Monto"] > 0)
].drop_duplicates(subset=["ID_Transaccion"])

# Agrupar total acumulado por Vendedor con la data limpia
df_resumen = df_limpio.groupby("Vendedor", as_index=False)["Monto"].sum()

# Guardar base limpia en Excel
nombre_excel = "Reporte_Auditado_Limpio.xlsx"
df_resumen.to_excel(nombre_excel, index=False)

# 3. FORMATO PROFESIONAL EN OPENPYXL
wb = load_workbook(nombre_excel)
ws = wb.active

fuente_encabezado = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
relleno_encabezado = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

for cell in ws[1]:
    cell.font = fuente_encabezado
    cell.fill = relleno_encabezado
    cell.alignment = Alignment(horizontal="center")

for row in range(2, ws.max_row + 1):
    ws.cell(row=row, column=2).number_format = '"$"#,##0.00'

for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

wb.save(nombre_excel)
print("¡Proceso finalizado! Se generó 'Reporte_Auditado_Limpio.xlsx' y el log 'auditoria_errores_detectados.csv'.")