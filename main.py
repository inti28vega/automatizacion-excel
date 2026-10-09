import os
import pandas as pd
from pipeline_modulos import (
    extraer_y_consolidar_datos,
    auditar_y_limpiar_datos,
    aplicar_formato_excel
)

def ejecutar_plantilla_maestra():
    print("🚀 Iniciando Plantilla Maestra de Automatización...")
    
    # 1. Extracción desde carpeta de origen
    directorio_origen = "reportes_anuales"
    df_crudo = extraer_y_consolidar_datos(directorio_origen)
    print(f"   [1/3] Extracción completada. Filas leídas: {len(df_crudo)}")
    
    # 2. Limpieza y Auditoría
    df_limpio, df_errores = auditar_y_limpiar_datos(df_crudo)
    
    if not df_errores.empty:
        df_errores.to_csv("log_errores_maestro.csv", index=False)
        print(f"   ⚠️ Se encontraron {len(df_errores)} registros con errores. Guardados en 'log_errores_maestro.csv'.")
    
    # 3. Consolidado e informe final
    df_resumen = df_limpio.groupby("Vendedor", as_index=False)["Monto"].sum()
    archivo_salida = "Reporte_Maestro_Final.xlsx"
    df_resumen.to_excel(archivo_salida, index=False)
    
    # Formato visual
    aplicar_formato_excel(archivo_salida)
    print(f"   [3/3] Reporte generado y formateado exitosamente en '{archivo_salida}'.")
    print("✅ ¡Pipeline completado con éxito!")

if __name__ == "__main__":
    ejecutar_plantilla_maestra()