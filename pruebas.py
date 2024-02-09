from datetime import date
import pandas as pd
import os
import shutil

archivo_root = "C:/Users/alexa/Documents/Manejo_Archivos"

for filename in os.listdir(archivo_root):
    if filename.endswith('.xlsx'):
        archivo = filename

data = pd.read_excel(archivo, sheet_name="Hoja 1", header=0, index_col="Connid")

df = pd.DataFrame(data, columns=['Poliza', 'Telefono',
                                 'ID_UNICO', 'Intento', 'Created', 'Calificacion',
                                 'Tipificacion', 'SubTipificacion'])

change_root = "C/Users/alexa/Documents/Manejo_Archivos/Archivos_higienizados"

fecha = date.today()

formateo = fecha.strftime('%Y%m%d')

columns = [
    ('PredictiveResultID', 1, True)
]

for columna, estatus, confirmacion in columns:
    data.drop(columna, axis = estatus, inplace = confirmacion)

df.to_csv('Base_117.csv')


for filename in os.listdir(archivo_root):
    if filename.endswith('.csv'):
        source_path = os.path.join(archivo_root, filename)
        target_path = os.path.join(change_root, filename)
        shutil.move(source_path, target_path)

for file in os.listdir(change_root):
    if filename.endswith('.csv'):
        file_path = os.path.join(change_root, file)
        name, ext = os.path.splitext(file)
        new_name = f'{name}_{formateo}{ext}'
        new_path = os.path.join(change_root, new_name)
        os.rename(file_path, new_path)

for file in os.listdir(archivo_root):
    if file.endswith('.xlsx'):
        os.remove(file)