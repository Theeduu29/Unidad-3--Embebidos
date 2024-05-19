import pandas as pd
import numpy as np
import glob

# Función para normalizar los datos
def normalizar(data):
    return (data - np.min(data)) / (np.max(data) - np.min(data))

# Ruta del directorio donde están los archivos Excel
ruta_archivos = 'C:\\embebidos\\Pot_Emb_Csv\\Forma2\\*.xlsx'

# Lista para almacenar los datos normalizados
dataframes_normalizados = []

# Leer cada archivo Excel, normalizar los datos y guardar en una lista hacia abajo
for archivo in glob.glob(ruta_archivos):
    df = pd.read_excel(archivo)
    # solo se puede llegar a usar si se quiere normalizar excel
    # Si solo quieres normalizar ciertas columnas, ajusta esta parte
    df_normalizado = df.apply(normalizar)
    dataframes_normalizados.append(df_normalizado)

# Genera con nombres los nuevos archivos de excel
nombres_nuevos_archivos = ['Promnormalizado.xlsx', 'Mednormalizada.xlsx', 'ValMaynormalizado.xlsx', 'ValMennormalizado.xlsx']

# Guarda todos los datos en el nuevo exel
for df, nombre in zip(dataframes_normalizados, nombres_nuevos_archivos):
    df.to_excel(nombre, index=False)

print("Archivos normalizados guardados exitosamente.")
