import pandas as pd

# 1. Simulación de datos "sucios" típicos de una PYME
datos = {
    "Fecha": ["2026-09-01", "2026-09-02", "2026-09-02", None, "2026-09-05"],
    "Concepto": ["  papelería ", "Internet ", "Mantenimiento", "Alquiler", "Suministros  "],
    "Monto": [45.0, 120.0, None, 500.0, 85.5],
    "Estado": ["Aprobado", "Aprobado", "Pendiente", "Aprobado", "Pendiente"]
}

df = pd.DataFrame(datos)

print("--- DATOS ORIGINALES ---")
print(df)
print("\n" + "="*40 + "\n")

# 2. Limpieza de texto (elimina espacios extras en los extremos y pone mayúscula inicial)
df["Concepto"] = df["Concepto"].str.strip().str.capitalize()

# 3. Tratamiento de valores nulos (rellena montos vacíos con 0 y fechas con '2026-09-01')
df["Monto"] = df["Monto"].fillna(0)
df["Fecha"] = df["Fecha"].fillna("2026-09-01")

# 4. Filtrado: Solo gastos Aprobados de más de $50
df_filtrado = df[(df["Estado"] == "Aprobado") & (df["Monto"] > 50)]

print("--- DATOS LIMPIOS Y FILTRADOS ---")
print(df_filtrado)

# 5. Guardar el reporte limpio en Excel
df_filtrado.to_excel("Reporte_Gastos_Limpio.xlsx", index=False)
print("\n¡Reporte limpio generado con éxito en 'Reporte_Gastos_Limpio.xlsx'!")