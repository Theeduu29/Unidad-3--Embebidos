import pandas as pd
import numpy as np
from pathlib import Path

# Función para estandarizar los datos
def estandarizacion(data):
    media = np.mean(data)
    desv_est = np.std(data)
    return (data - media) / desv_est

# Rutas de los archivos Excel que se van a estandarizar
rutas_archivos = [
    Path('C:/embebidos/Pot_Emb_Csv/Forma2/Promedio.xlsx'),
    Path('C:/embebidos/Pot_Emb_Csv/Forma2/Mediana.xlsx'),
    Path('C:/embebidos/Pot_Emb_Csv/Forma2/ValorMayor.xlsx'),
    Path('C:/embebidos/Pot_Emb_Csv/Forma2/ValorMenor.xlsx'),
]

# Lista para almacenar los DataFrames estandarizados
dataframes_estandarizados = []

# Leer cada archivo Excel, estandarizar los datos y guardar en una lista
for ruta_archivo in rutas_archivos:
    try:
        df = pd.read_excel(ruta_archivo)
        # Seleccionar solo columnas numéricas para la estandarización
        df_estandarizado = df.select_dtypes(include=[np.number]).apply(estandarizacion)
        # Conservar las columnas no numéricas sin cambios
        df_estandarizado = df_estandarizado.join(df.select_dtypes(exclude=[np.number]))
        dataframes_estandarizados.append(df_estandarizado)
    except Exception as e:
        print(f"Error al procesar el archivo {ruta_archivo}: {e}")

# Generar nombres para los nuevos archivos Excel
nombres_nuevos_archivos = [
    'C:/embebidos/Pot_Emb_Csv/Forma2/Promestandarizado.xlsx',
    'C:/embebidos/Pot_Emb_Csv/Forma2/Medestandarizada.xlsx',
    'C:/embebidos/Pot_Emb_Csv/Forma2/ValMayestandarizado.xlsx',
    'C:/embebidos/Pot_Emb_Csv/Forma2/ValMenestandarizado.xlsx',
]

# Guardar cada DataFrame estandarizado en un nuevo archivo Excel
for df, nombre in zip(dataframes_estandarizados, nombres_nuevos_archivos):
    try:
        df.to_excel(nombre, index=False)
    except Exception as e:
        print(f"Error al guardar el archivo {nombre}: {e}")

print("Archivos estandarizados guardados exitosamente.")
