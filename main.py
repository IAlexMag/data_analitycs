from datetime import date
import pandas as pd
import os
import shutil
#--------------------------limpieza de datos----------------------------------------------
#------variables---------------------
archivos_root = "C:/Users/alexa/Documents/Manejo_Archivos"

for filename in os.listdir(archivos_root):
    if filename.endswith('.xlsx'):
        archivo = filename

data = pd.read_excel(archivo, sheet_name="00 Retencion", header=6, index_col="Connid")
#data_1 = pd.read_excel(archivo, sheet_name="00 Retencion", header=6)

#data_2 = pd.read_excel(archivo, sheet_name="Hoja1")

#data_2.rename(columns={'ID_UNICO': 'Id unico'}, inplace = True)

#data = pd.merge(data_1, data_2, on = 'Id unico', index_col = 'Connid')

df = pd.DataFrame(data, columns = ['Fecha', 'DN', 'Intento', 'Id unico','Username', 'Agente', 'Tipo_base', 'Calificacion', 'Tipificacion', 'SubTipificacion'])

change_root = "C:/Users/alexa/Documents/Manejo_Archivos/Archivos_higienizados"

fecha = date.today()

formateo = fecha.strftime('%Y%m%d')

columnas = [
    ('Campaña', 1, True),
    ('Campaña Desc', 1, True),
    ('Tipo Llamada', 1, True),
    ('TipoContacto', 1, True),
    ('Duration', 1, True),
    ('Talktime', 1, True),
    ('Modo de marcacion', 1, True),
    ('Id', 1, True),
    ('Apaterno', 1, True),
    ('Amaterno', 1, True),
    ('Nombres', 1, True),
    ('Celular', 1, True),
    ('RFC', 1, True),
    ('Poliza', 1, True),
    ('Correo', 1, True),
    ('CentroServicio', 1, True),
    ('Contratante', 1, True),
    ('DCN', 1, True),
    ('RollId', 1, True),
    ('Siniestro', 1, True),
    ('Notas', 1, True)
]
#-----------------------------------------------------------------------------------------

#-----------------------Inicia limpieza de archivo excel--------------------------------------------------------------------------------------
for columna, estatus, confirmacion in columnas:
    data.drop(columna, axis = estatus, inplace = confirmacion)
#print(data.columns)
#print(data.head(20))

df.to_csv('Higienizado.csv')

#-------------------------Termina limpieza de archivo excel-------------------------------------------------------------------------------------

#------------------------Inicia movimiento y renombrado de directorios--------------------------------------------------------------------------

#-----mueve el csv a la carpeta Archivos_higienizados---------------
for filename in os.listdir(archivos_root):
    if filename.endswith('.csv'):
        source_path = os.path.join(archivos_root, filename)
        target_path = os.path.join(change_root, filename)
        shutil.move(source_path, target_path)
#------------------------------------------------------------------

#-----renombra el csv y agrega la fecha actual---------------------
for file in os.listdir(change_root):
    if file.endswith('.csv'):
        file_path = os.path.join(change_root, file)
        name, ext = os.path.splitext(file)
        new_name = f'{name}_{formateo}{ext}'
        new_path = os.path.join(change_root, new_name)
        os.rename(file_path, new_path)
#-------------------------------------------------------------------

#----------elimina el archivo .xlsx---------------------------------
for file in os.listdir(archivos_root):
    if file.endswith('.xlsx'):
        os.remove(file)
#------------------------------------------------------------------------------------------------------------------------------------------