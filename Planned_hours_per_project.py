#Planned hours per project
import pandas as pd
import openpyxl
import pylab

workbook_assignament_matrix = 'C:/Users/User/Documents/BPS Technology Solutions/Matriz Asignacion TD_v1.xlsx' # Import the Excel archive
dataframe_pandas = pd.read_excel(workbook_assignament_matrix) # convert archive in dataframe - Pandas
#print(dataframe_pandas)
#print(dataframe_pandas['Recurso'])
#valor_celda_persona = dataframe_pandas.loc[166, 'Recurso']
#print(valor_celda_persona)
searched_name = "Paola Saboyá"
index = 1

for index, name in dataframe_pandas['Recurso'].items():
    #print(name)
    if name == searched_name:
       print(f"El nombre del recurso se encuentra en la fila {index}")
       break
else:
    print("El nombre del recurso no está en la matriz")
       