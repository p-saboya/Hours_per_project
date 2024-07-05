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

#Function --> Search resource's name to find the Excel row of the first and the last projects assigned to he/she
#The result is de indexes of the Excel row first project and Excel row final project, 
#Excel row final project is calculate with the sume of total of projects assigned to the person
def find_excel_projects_location() :
    searched_name = "Paola Saboyá"
    #search_name = input("enter resource's name ")
    index = 1
    for index, name in dataframe_pandas['Recurso'].items() : 
        #print(name)
        if name == searched_name :
           #print(f"The resource name is in the row {index}")
           Excel_row_first_project = index
           break
    else :
        print("Resource name is not in array")
    invalid_project_number = True
    while invalid_project_number == True :
        number_total_assigned_projects = input("enter total of assigned projects => ")
        if not number_total_assigned_projects.isnumeric() or number_total_assigned_projects == "0" :
            print('Invalid option') 
        else:
            invalid_project_number = False
    Excel_row_final_project = Excel_row_first_project + int(number_total_assigned_projects)
    return Excel_row_first_project, Excel_row_final_project, number_total_assigned_projects
    
Excel_row_first_project, Excel_row_final_project, number_total_assigned_projects = find_excel_projects_location()
#print(type(Excel_row_first_project))

#Funtion --> to validate that the days of the week are between 1 and 5
def receiving_number_working_days_converts_to_hours() :
    list_number_days_week = []
    for day in range(1, 6):
        list_number_days_week.append(str(day))
    #print(list_number_days_week)
    invalid_number_working_days = True
    while invalid_number_working_days == True :
        number_working_days_this_week = input("enter working days this week => ")
        if not number_working_days_this_week in list_number_days_week :
            print('Invalid option')
        else:
            invalid_number_working_days = False
    number_working_hours_this_week = int(number_working_days_this_week)*8 # 8 h is the number of working hours per day in Colombia
    print("The number of working hours this week is => ", number_working_hours_this_week, "h")
    return number_working_hours_this_week
    
number_working_hours_this_week = receiving_number_working_days_converts_to_hours()

#Function --> It creates the arrangement of dedication percentages for each project for the week
def arragement_calculated_percentage_calculation_this_week(number_total_assigned_projects, number_working_hours_this_week, dataframe_pandas) : 
    cumulative_of_hours = []
    arregement_percentage_calculation_this_week = []
    sum_arregement_percentage_calculation_this_week = 0
    while sum_arregement_percentage_calculation_this_week < 100 :
        for project in range(Excel_row_first_project, Excel_row_final_project, 1) :
            project_name = dataframe_pandas.loc[project, 'Descripcion']
            print(project_name)
            #While --> to choose the project weight (in hours) during the following week
            invalid_weight = True
            while invalid_weight == True:
              hours_proj_this_week = input("enter number of hours per week estimated for each project => ")
              if not hours_proj_this_week.isnumeric() or int(hours_proj_this_week) < 0 or int(hours_proj_this_week) > number_working_hours_this_week :
                  print('Invalid option')
              else:
                  invalid_weight = False
                  hours_proj_this_week = int(hours_proj_this_week)
            cumulative_of_hours.append(hours_proj_this_week)
            cumulative_sum_of_hours = sum(cumulative_of_hours)
            print("cumulative sum of hours => ", cumulative_sum_of_hours)
            percentage_calculation_this_week = (hours_proj_this_week * 100) / number_working_hours_this_week
            print("percentage_calculation_this_week => % ", percentage_calculation_this_week)
            arregement_percentage_calculation_this_week.append(percentage_calculation_this_week)
        print("arregement_percentage_calculation_this_week ",arregement_percentage_calculation_this_week)
        sum_arregement_percentage_calculation_this_week = sum(arregement_percentage_calculation_this_week)
        print("sum_arregement_percentage_calculation_this_week => % ", sum_arregement_percentage_calculation_this_week)
    return arregement_percentage_calculation_this_week

arregement_percentage_calculation_this_week = arragement_calculated_percentage_calculation_this_week(number_total_assigned_projects, number_working_hours_this_week, dataframe_pandas)

#Funtion --> It modifies the Excel file, writes the arregement_percentage_calculation_this_week on the column of the chose month 
def modifinding_dataframe_pandas(dataframe_pandas, arregement_percentage_calculation_this_week, Excel_row_first_project, Excel_row_final_project) :
    range_Excel_rows_to_modify = range(Excel_row_first_project, Excel_row_final_project)
    #print("Excel_row_first_project => ", Excel_row_first_project)
    #print("Excel_row_final_project =>", Excel_row_final_project)
    dataframe_pandas.loc[range_Excel_rows_to_modify, 'Julio'] = arregement_percentage_calculation_this_week
    dataframe_pandas.to_excel('C:/Users/User/Documents/BPS Technology Solutions/Matriz Asignacion TD_v2.xlsx', index=True, header=True)
    return None, None
    
def main() :
    modifinding_dataframe_pandas(dataframe_pandas, arregement_percentage_calculation_this_week, Excel_row_first_project, Excel_row_final_project) 
    return None, None
    
if __name__ == "__main__" :
    main()

