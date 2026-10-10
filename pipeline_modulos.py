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

def auditar_y_limpiar_datos(df, col_vendedor, col_monto):
    """Limpia textos, aisla errores y retorna DataFrames usando columnas dinámicas."""
    df_temp = df.copy()
    
    if col_vendedor in df_temp.columns:
        df_temp[col_vendedor] = df_temp[col_vendedor].astype(str).str.strip()
    
    nulos = df_temp[df_temp[col_vendedor] == "None"] if col_vendedor in df_temp.columns else pd.DataFrame()
    montos_invalidos = df_temp[df_temp[col_monto] <= 0] if col_monto in df_temp.columns else pd.DataFrame()
    
    df_errores = pd.concat([nulos, montos_invalidos]).drop_duplicates()
    
    df_limpio = df_temp[
        (df_temp[col_vendedor] != "None") & 
        (df_temp[col_monto] > 0)
    ].drop_duplicates() if col_monto in df_temp.columns else df_temp
    
    return df_limpio, df_errores

def aplicar_formato_excel(archivo_salida, formato_moneda):
    """Aplica estilos profesionales y el formato de moneda configurado."""
    wb = load_workbook(archivo_salida)
    ws = wb.active

    fuente_encabezado = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    relleno_encabezado = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    for cell in ws[1]:
        cell.font = fuente_encabezado
        cell.fill = relleno_encabezado
        cell.alignment = Alignment(horizontal="center")

    for row in range(2, ws.max_row + 1):
        ws.cell(row=row, column=2).number_format = formato_moneda

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    wb.save(archivo_salida)