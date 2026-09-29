import os
import pandas as pd

ruta = r"C:\Users\user\Desktop\gastos"
archivos = [f for f in os.listdir(ruta) if f.endswith(".xlsx") and not f.startswith("~$")]

lista_df = []

for archivo in archivos:
    path_completo = os.path.join(ruta, archivo)
    df = pd.read_excel(path_completo)
    
    # Agrega la columna sucursal sin la extensión .xlsx
    df["Sucursal"] = archivo.replace(".xlsx", "")
    lista_df.append(df)

# Une todos los DataFrames de la lista en uno solo
df_consolidado = pd.concat(lista_df, ignore_index=True)

# Guarda el resultado en un archivo Excel nuevo
df_consolidado.to_excel("Consolidado_Python.xlsx", index=False)

print(f"Listo: se consolidaron {len(df_consolidado)} filas de {len(archivos)} sucursales.")