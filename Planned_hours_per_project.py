#Planned hours per project
import pandas as pd
import openpyxl
import pylab
import calendar
import locale
locale.setlocale(locale.LC_ALL, 'es_ES.UTF-8') #Spahinsh configuration for the name of the months

excel_assignament_matrix = 'C:/Users/User/Documents/BPS Technology Solutions/Matriz Asignacion TD_v2.xlsx' # Import the Excel archive
dataframe_pandas = pd.read_excel(excel_assignament_matrix) # convert archive in dataframe - Pandas
#assinement_sheet = dataframe_pandas['Asignacion']
#print(dataframe_pandas)
#print(dataframe_pandas['Recurso'])
#valor_celda_persona = dataframe_pandas.loc[166, 'Recurso']
#print(valor_celda_persona)

'''
#Function --> Search month and week to put the assinated percentage
def find_month_week_location() :
    print("dentro de find month")
    list_months = calendar.month_name[1:]
    #print(list_months)
    invalid_month = True
    while invalid_month == True :
        searched_month = input("enter the month of the year that you want to modify => ")
        if not searched_month in list_months :
            print('Invalid option')
        else:
            invalid_month = False
    print(searched_month)
    firts_row = dataframe_pandas.head(1)
    print(firts_row)
    print(type(firts_row))
    for month in firts_row.items() :
        #valor_celda = month.value
        if valor_celda == searched_month :
            print(f"Dato encontrado en la columna: {month.column}")
            break
    return searched_month

searched_month = find_month_week_location()'''


#Function --> Search resource's name to find the Excel row of the first and the last projects assigned to he/she
#The result is de indexes of the Excel row first project and Excel row final project, 
#Excel row final project is calculate with the sume of total of projects assigned to the person
def find_excel_projects_location() :
    #searched_name = input("enter the talent name that you want to find => ")
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

#Funtion --> This is to validate that the days of the week are between 1 and 5
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

#Function --> This calculate the percentage of each project per week (To divide the While to improve the second time calculation, and only ask for the projects with time different to zero)
def percentage_calculation_this_week_f(hours_proj_this_week) :
    percentage_calculation_this_week = (hours_proj_this_week * 100) / number_working_hours_this_week
    percentage_calculation_this_week_round = round(percentage_calculation_this_week, 3)
    print("percentage_calculation_this_week => % ", percentage_calculation_this_week_round)
    return percentage_calculation_this_week_round

#Function --> This creates the arrangement of dedication percentages for each project for the week
def arragement_calculated_percentage_calculation_this_week(number_total_assigned_projects, number_working_hours_this_week, dataframe_pandas) : 
    cumulative_of_hours = []
    cumulative_of_percentage = []
    arregement_percentage_calculation_this_week = []
    arregement_percentage_calculation_this_week_str = []
    sum_arregement_percentage_calculation_this_week = 0 
    manual_approval = False
    while manual_approval == False :
        for project in range(Excel_row_first_project, Excel_row_final_project, 1) :
            project_name = dataframe_pandas.loc[project, 'Descripcion']
            print('----------------')
            print(project_name)
            invalid_weight = True
            while invalid_weight == True:
              hours_proj_this_week = input("enter number of hours per week estimated for each project => ")
              if not hours_proj_this_week.isnumeric() or int(hours_proj_this_week) < 0 or int(hours_proj_this_week) > number_working_hours_this_week :
                  print('Invalid option')
              else:
                  invalid_weight = False
                  hours_proj_this_week = int(hours_proj_this_week)
            #Print cumulative of hours
            cumulative_of_hours.append(hours_proj_this_week)
            cumulative_sum_of_hours = sum(cumulative_of_hours)
            print("cumulative sum of hours => ", cumulative_sum_of_hours)
            #Call to percentage calculation function
            percentage_calculation_this_week_round = percentage_calculation_this_week_f(hours_proj_this_week)
            cumulative_of_percentage.append(percentage_calculation_this_week_round)
            cumulative_sum_of_percentage_this_week = sum(cumulative_of_percentage)
            print("cumulative_sum_of_percentage_this_week => % ", cumulative_sum_of_percentage_this_week)            
            arregement_percentage_calculation_this_week.append(percentage_calculation_this_week_round)
            arregement_percentage_calculation_this_week_str.append(str(percentage_calculation_this_week_round) + " %")
        print("arregement_percentage_calculation_this_week_str ",arregement_percentage_calculation_this_week_str)
        sum_arregement_percentage_calculation_this_week = sum(arregement_percentage_calculation_this_week)
        print("sum_arregement_percentage_calculation_this_week => % ", sum_arregement_percentage_calculation_this_week)
        
        manual_approval = input("Do you agree with the accumulated sum for the week of the assigned percentages for the projects?, If your answer is yes, please write Y if not, please write N => " )
        manual_approval = manual_approval.lower()
        if manual_approval == "y" : 
            manual_approval = True
        else :
            manual_approval = False
            arregement_percentage_calculation_this_week = [0] * len(arregement_percentage_calculation_this_week)
            #print(arregement_percentage_calculation_this_week)
            #cumulative_of_hours = [0] * len(cumulative_of_hours)4
            cumulative_of_hours = []
            cumulative_sum_of_hours = 0
            print('*** NEW HOURS PER WEEK ESTIMATION***')
        #print("final manual_approval", manual_approval)
        #print(type(manual_approval))
    return arregement_percentage_calculation_this_week_str
    
arregement_percentage_calculation_this_week_str = arragement_calculated_percentage_calculation_this_week(number_total_assigned_projects, number_working_hours_this_week, dataframe_pandas)

'''def cumulative_sum_of_hours_and_percentage(hours_proj_this_week) :
    cumulative_of_hours.append(hours_proj_this_week)
    cumulative_sum_of_hours = sum(cumulative_of_hours)
    print("cumulative sum of hours => ", cumulative_sum_of_hours)
    return None, None'''

#Funtion --> This modifies the Excel file, writes the arregement_percentage_calculation_this_week on the column of the chose month 
def modifinding_dataframe_pandas(dataframe_pandas, arregement_percentage_calculation_this_week_str, Excel_row_first_project, Excel_row_final_project) :
    range_Excel_rows_to_modify = range(Excel_row_first_project, Excel_row_final_project)
    dataframe_pandas.loc[range_Excel_rows_to_modify, 'Febrero 2025'] = arregement_percentage_calculation_this_week_str
    dataframe_pandas.to_excel('C:/Users/User/Documents/BPS Technology Solutions/Matriz Asignacion TD_v2.xlsx', index=True, header=True)
    #dataframe_pandas.to_excel('C:/Users/User/Documents/BPS Technology Solutions/Matriz Asignacion TD_v2.xlsx', index=True,  style=lambda x: x.apply(lambda v: v.rjust(len(v), '.') + '%'))
    return None, None
    
def main() :
    modifinding_dataframe_pandas(dataframe_pandas, arregement_percentage_calculation_this_week_str, Excel_row_first_project, Excel_row_final_project) 
    return None, None
    
if __name__ == "__main__" :
    main()
