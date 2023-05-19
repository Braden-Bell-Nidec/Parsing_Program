'''
Braden Bell
05-19-23
Description:
This Python script is designed to perform an analysis of job responsibilities within various departments of an organization. 
The user is first asked to input a file name and an outlier percentage threshold, which assists in identifying responsibilities 
that occur less frequently, or outliers. 
'''

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment

#Adjusts the column width of all specified columns in the sheet
def adjust_column_width(sheet, cols_width_dict):
    for col, width in cols_width_dict.items():
        sheet.column_dimensions[col].width = width

#Aligns the cells in the specified columns to the given type (center, left, right)
def align_cells(sheet, cols, alignment):
    for col in cols:
        for cell in sheet[col]:
            cell.alignment = alignment


#Generates a pie chart in the specified sheet based on responsibility data
def create_pie_chart(sheet, responsibilities):
    chart = PieChart()
    labels = Reference(sheet, min_col=1, min_row=2, max_row=len(responsibilities)+1)
    data = Reference(sheet, min_col=2, min_row=1, max_row=len(responsibilities)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = 'Responsibility Distribution'
    sheet.add_chart(chart, "C1")

#Appends all the rows of a DataFrame to a given sheet object
def append_dataframe_to_sheet(sheet, df):
    for row in dataframe_to_rows(df, index=False, header=True):
        sheet.append(row)


#Retrieves responsibility data from the given DataFrame and returns it
def get_responsibility_data(df):
    responsibilities = df['RESPONSIBILITY_NAME'].value_counts().reset_index()
    responsibilities.columns = ['RESPONSIBILITY_NAME', 'COUNTS']
    return responsibilities

#Uses create_pie_chart() to write the pie charts to the sheet. It contains formatting data.
def create_pie_charts(df, wb, sheetname):
    ws = wb[sheetname]
    responsibilities = get_responsibility_data(df)
    append_dataframe_to_sheet(ws, responsibilities)
    adjust_column_width(ws, {'A': 45, 'B': 10})
    align_cells(ws, ['B'], Alignment(horizontal='center'))
    create_pie_chart(ws, responsibilities)

#Set terminal output options
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', None)

#Get file name from user
fileName = input("Enter file name or path: ") 
if not fileName.endswith(".xlsx"):
    fileName += (".xlsx")
    
#Attempt to open excel sheet
try:
    df = pd.read_excel(fileName)
except:
    print("Error reading file!")
    input("Nothing to do! Press enter to close.")
    exit()

#Get custom percentage from user if needed
userPercent = input("Enter outlier percentage threshold (default is 7%): ")
try:
    if userPercent == "":
        userPercent = 0.07
    else:
        userPercent = abs(userPercent / 100.00)
except:
    print("Input value error, resorting to default")
    userPercent = 0.07

# Create a dictionary of DataFrames for each unique department
department_dfs = {}
for department in df['DEPARTMENT'].drop_duplicates().values:
    department_dfs[department] = df[df['DEPARTMENT'] == department]


counts = df.groupby(['DEPARTMENT', 'JOB_TITLE', 'RESPONSIBILITY_NAME']).size().reset_index(name='counts')
total_counts = counts.groupby(['DEPARTMENT', 'JOB_TITLE'])['counts'].sum().reset_index(name='total_counts')
counts = pd.merge(counts, total_counts, on=['DEPARTMENT', 'JOB_TITLE'])
counts['percentage'] = counts['counts'] / counts['total_counts']

outliers = counts[counts['percentage'] < userPercent] #Calculate the outliers

#Set up parameters for outlier dataframe
outlier_users_list = [
    {
        'USER_NAME': user, 
        'DEPARTMENT': outlier['DEPARTMENT'],
        'JOB_TITLE': outlier['JOB_TITLE'], 
        'RESPONSIBILITY_NAME': outlier['RESPONSIBILITY_NAME'],
        'PERCENTAGE': round(outlier['percentage']*100,2)
    }
    for _, outlier in outliers.iterrows() 
    for user in df[
        (df['DEPARTMENT'] == outlier['DEPARTMENT']) & 
        (df['JOB_TITLE'] == outlier['JOB_TITLE']) & 
        (df['RESPONSIBILITY_NAME'] == outlier['RESPONSIBILITY_NAME'])
    ]['USER_NAME']
]

outlier_users = pd.DataFrame(outlier_users_list)
print("\n\n-=================================== Possible Outliers ===================================-")
print(outlier_users) #Print outlier data to terminal

#Set up the paremeters for the non-outlier DataFrame
non_outlier_list = [
    {
        'DEPARTMENT': row['DEPARTMENT'],
        'JOB_TITLE': row['JOB_TITLE'], 
        'RESPONSIBILITY_NAME': row['RESPONSIBILITY_NAME'],
        'PERCENTAGE': round(row['percentage']*100,2)
    }
    for _, row in counts.iterrows()
    if not (
        (outliers['DEPARTMENT'] == row['DEPARTMENT']) & 
        (outliers['JOB_TITLE'] == row['JOB_TITLE']) & 
        (outliers['RESPONSIBILITY_NAME'] == row['RESPONSIBILITY_NAME'])
    ).any()
]

non_outliers = pd.DataFrame(non_outlier_list)
print("\n\n-==================== Non-Outliers ====================-")
print(non_outliers) #Print non-outlier DataFrame data to terminal

wb = openpyxl.Workbook() #Open workbook
wb.remove(wb.active)  #Remove the default sheet

#Create and populate the outlier sheet
outliers_sheet = wb.create_sheet(title='Outliers')
append_dataframe_to_sheet(outliers_sheet, outlier_users)

#Format the outlier sheet
adjust_column_width(outliers_sheet, {'A': 15, 'B': 15, 'C': 35, 'D': 42, 'E': 12})
align_cells(outliers_sheet, ['A', 'B', 'C', 'D'], Alignment(horizontal='center'))

#Create and populate the non-outlier sheet
non_outliers_sheet = wb.create_sheet(title='Non-Outliers')
append_dataframe_to_sheet(non_outliers_sheet, non_outliers)

#Format non-outlier sheet
adjust_column_width(non_outliers_sheet, {'A': 15, 'B': 35, 'C': 32, 'D': 12})
align_cells(non_outliers_sheet, ['A', 'B', 'C', 'D'], Alignment(horizontal='center'))


for department, df in department_dfs.items():
    sheetname = department[:31]  #Sheet names can't be longer than 31 characters or Excel breaks
    wb.create_sheet(title=sheetname)
    create_pie_charts(df, wb, sheetname)

#Save changes to the sheet
wb.save("analysis.xlsx")

input("\n\nAll tasks completed. Press enter to close.")

