import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

def extraer_y_consolidar_datos(directorio_base):
    """Recorre carpetas anidadas y consolida todos los Excel/CSV encontrados."""
    registros = []
    for root, _, files in os.walk(directorio_base):
        for file in files:
            ruta_completa = os.path.join(root, file)
            if file.endswith(".xlsx") or file.endswith(".xls"):
                hojas = pd.read_excel(ruta_completa, sheet_name=None)
                for _, df in hojas.items():
                    registros.append(df)
            elif file.endswith(".csv"):
                df = pd.read_csv(ruta_completa)
                registros.append(df)
    
    if registros:
        return pd.concat(registros, ignore_index=True)
    return pd.DataFrame()

def auditar_y_limpiar_datos(df):
    """Limpia textos, aisla errores y retorna DataFrames separados (Limpio y Errores)."""
    df_temp = df.copy()
    
    # Normalización de textos
    if "Vendedor" in df_temp.columns:
        df_temp["Vendedor"] = df_temp["Vendedor"].astype(str).str.strip()
    
    # Identificación de anomalías
    nulos = df_temp[df_temp["Vendedor"] == "None"] if "Vendedor" in df_temp.columns else pd.DataFrame()
    montos_invalidos = df_temp[df_temp["Monto"] <= 0] if "Monto" in df_temp.columns else pd.DataFrame()
    
    df_errores = pd.concat([nulos, montos_invalidos]).drop_duplicates()
    
    # Data limpia
    df_limpio = df_temp[
        (df_temp["Vendedor"] != "None") & 
        (df_temp["Monto"] > 0)
    ].drop_duplicates() if "Monto" in df_temp.columns else df_temp
    
    return df_limpio, df_errores

def aplicar_formato_excel(archivo_salida):
    """Aplica estilos profesionales al archivo Excel generado."""
    wb = load_workbook(archivo_salida)
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

    wb.save(archivo_salida)