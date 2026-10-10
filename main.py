import json
import os
import pandas as pd
from pipeline_modulos import (
    extraer_y_consolidar_datos,
    auditar_y_limpiar_datos,
    aplicar_formato_excel
)

def cargar_configuracion(ruta_config="config.json"):
    with open(ruta_config, "r", encoding="utf-8") as f:
        return json.load(f)

def ejecutar_plantilla_maestra():
    print("🚀 Cargando archivo de configuración 'config.json'...")
    config = cargar_configuracion()
    
    print("🚀 Iniciando Plantilla Maestra Parametrizada...")
    
    # 1. Extracción
    df_crudo = extraer_y_consolidar_datos(config["directorio_origen"])
    print(f"   [1/3] Extracción completada. Filas leídas: {len(df_crudo)}")
    
    # 2. Limpieza y Auditoría con parámetros dinámicos
    df_limpio, df_errores = auditar_y_limpiar_datos(
        df_crudo, 
        config["columna_vendedor"], 
        config["columna_monto"]
    )
    
    if not df_errores.empty:
        df_errores.to_csv(config["log_errores"], index=False)
        print(f"   ⚠️ Registros con errores detectados. Log generado en '{config['log_errores']}'.")
    
    # 3. Consolidación e informe final
    df_resumen = df_limpio.groupby(config["columna_vendedor"], as_index=False)[config["columna_monto"]].sum()
    df_resumen.to_excel(config["archivo_salida"], index=False)
    
    # Formato visual dinámico
    aplicar_formato_excel(config["archivo_salida"], config["moneda_formato"])
    print(f"   [3/3] Reporte generado exitosamente en '{config['archivo_salida']}'.")
    print("✅ ¡Pipeline parametrizado completado con éxito!")

if __name__ == "__main__":
    ejecutar_plantilla_maestra()