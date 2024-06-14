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

#For --> Search resource's name to find the Excel row of the first project assigned to he/she
searched_name = "Paola Saboyá"
index = 1
for index, name in dataframe_pandas['Recurso'].items() : 
    #print(name)
    if name == searched_name :
       print(f"The resource name is in the row {index}")
       Excel_row_first_project = index
       break
else :
    print("Resource name is not in array")

number_total_assigned_projects = input("enter total of assigned projects ")
number_working_days_this_week = input("enter working days this week ")
number_working_hours_this_week = int(number_working_days_this_week)*8 # 8 is the number of working hours per day
Excel_row_final_project = Excel_row_first_project + int(number_total_assigned_projects)
print(Excel_row_final_project)
#final_project_name = dataframe_pandas.loc[Excel_row_final_project, 'Descripcion']
#print(final_project_name)

#Excel_row_first_project = 1 #Test
#Excel_row_final_project = 2 + 1 #Test
#For --> Find the projects assigned to the resource one by one, then it calculates the percentage of dedication according to priority and working hours of the week
for project in range(Excel_row_first_project, Excel_row_final_project, 1) :
    project_name = dataframe_pandas.loc[project, 'Descripcion']
    print(project_name)
    weight_proj_this_week = input("enter project weight (number) : 0 => No Started or On hold, 1 => Minimum effort, 2 => Average effort, 3 => Demanding ")
    weight_proj_this_week = int(weight_proj_this_week)

    if weight_proj_this_week == 0 :
        number_hours_this_week = 0
        print("number_hours_this_week => ", number_hours_this_week)
    elif weight_proj_this_week == 1 :
        number_hours_this_week = 2
        print("number_hours_this_week => ", number_hours_this_week)
    elif weight_proj_this_week == 2 :
        number_hours_this_week = 4
        print("number_hours_this_week => ", number_hours_this_week)
    else :
        number_hours_this_week = 8
        print("number_hours_this_week => ", number_hours_this_week)

    percentage_calculation_this_week = (number_hours_this_week * 100) / number_working_hours_this_week
    print("percentage_calculation_this_week => ", percentage_calculation_this_week)

    #print(f"project number {project}")  
    dataframe_pandas.loc[project, "Junio"] = percentage_calculation_this_week
#print(dataframe_pandas['Febrero'])
dataframe_pandas.to_excel('C:/Users/User/Documents/BPS Technology Solutions/Matriz Asignacion TD_v2.xlsx', index=True, header=True)
