'''
Braden Bell
05-22-23
Description:
This Python script is designed to perform an analysis of job responsibilities within various departments of an organization. 
The user is first asked to input a file name and an outlier percentage threshold, which assists in identifying responsibilities 
that occur less frequently, or outliers. 
'''

import pandas as pd
import os
from Combiner import merge_files
import Excel_Functions as ef

DEFAULT_PERCENT = 0.07  # The default outlier percentage threshold

# Set terminal output options
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', None)

# Get file names from user
EPGA_File = input("Enter path of EPGA file: ")
if not EPGA_File.endswith('.xlsx'):
    EPGA_File += '.xlsx'

AD_File = input("Enter path of Active Directory file: ")
if not AD_File.endswith('.csv'):
    AD_File += '.csv'

fileName = merge_files(EPGA_File, AD_File, 'combined.xlsx') 

# Attempt to open excel sheet
try:
    df = pd.read_excel(fileName)
except:
    print("Error reading file!")
    input("Nothing to do! Press enter to close.")
    exit()

# Get outlier percentage from user
userPercent = input("Enter outlier percentage threshold (default is 7%): ")
if userPercent == "":
    userPercent = DEFAULT_PERCENT
else:
    try:
        userPercent = abs(float(userPercent) / 100.00)
    except:
        print("Input value error, resorting to default")
        userPercent = DEFAULT_PERCENT

# Create a dictionary of DataFrames for each unique department
department_dfs = {}
for department in df['DEPARTMENT'].drop_duplicates().values:
    department_dfs[department] = df[df['DEPARTMENT'] == department]

# Calculate counts and percentages
counts = df.groupby(['DEPARTMENT', 'JOB_TITLE', 'RESPONSIBILITY_NAME']).size().reset_index(name='counts')
total_counts = counts.groupby(['DEPARTMENT', 'JOB_TITLE'])['counts'].sum().reset_index(name='total_counts')
counts = pd.merge(counts, total_counts, on=['DEPARTMENT', 'JOB_TITLE'])
counts['percentage'] = counts['counts'] / counts['total_counts']

outliers = counts[counts['percentage'] < userPercent]  # Calculate the outliers

# Prepare outlier data
outlier_users_list = []
for _, outlier in outliers.iterrows():
    for user in df[
        (df['DEPARTMENT'] == outlier['DEPARTMENT']) & 
        (df['JOB_TITLE'] == outlier['JOB_TITLE']) & 
        (df['RESPONSIBILITY_NAME'] == outlier['RESPONSIBILITY_NAME'])
    ]['USER_NAME']:
        outlier_users_list.append({
            'USER_NAME': user, 
            'DEPARTMENT': outlier['DEPARTMENT'],
            'JOB_TITLE': outlier['JOB_TITLE'], 
            'RESPONSIBILITY_NAME': outlier['RESPONSIBILITY_NAME'],
            'PERCENTAGE': round(outlier['percentage']*100,2)
        })

outlier_users = pd.DataFrame(outlier_users_list)
print("\n\n-=================================== Possible Outliers ===================================-")
print(outlier_users)  # Print outlier data to terminal

# Prepare non-outlier data
non_outlier_list = []
for _, row in counts.iterrows():
    if not (
        (outliers['DEPARTMENT'] == row['DEPARTMENT']) & 
        (outliers['JOB_TITLE'] == row['JOB_TITLE']) & 
        (outliers['RESPONSIBILITY_NAME'] == row['RESPONSIBILITY_NAME'])
    ).any():
        non_outlier_list.append({
            'DEPARTMENT': row['DEPARTMENT'],
            'JOB_TITLE': row['JOB_TITLE'], 
            'RESPONSIBILITY_NAME': row['RESPONSIBILITY_NAME'],
            'PERCENTAGE': round(row['percentage']*100,2)
        })

non_outliers = pd.DataFrame(non_outlier_list)
print("\n\n-==================== Non-Outliers ====================-")
print(non_outliers)  # Print non-outlier DataFrame data to terminal

# Prepare to write to excel
wb = ef.Workbook()  # Open workbook
wb.remove(wb.active)  # Remove the default sheet

# Create and populate the outlier sheet
outliers_sheet = wb.create_sheet(title='Outliers')
ef.append_dataframe_to_sheet(outliers_sheet, outlier_users)

# Format the outlier sheet
ef.adjust_column_width(outliers_sheet, {'A': 15, 'B': 15, 'C': 35, 'D': 42, 'E': 12})
ef.align_cells(outliers_sheet, ['A', 'B', 'C', 'D'], ef.Alignment(horizontal='center'))

# Create and populate the non-outlier sheet
non_outliers_sheet = wb.create_sheet(title='Non-Outliers')
ef.append_dataframe_to_sheet(non_outliers_sheet, non_outliers)

# Format non-outlier sheet
ef.adjust_column_width(non_outliers_sheet, {'A': 15, 'B': 35, 'C': 32, 'D': 12})
ef.align_cells(non_outliers_sheet, ['A', 'B', 'C', 'D'], ef.Alignment(horizontal='center'))

invalid_chars = ['/', '\\', '?', '*', '[', ']', ':']

for department, df in department_dfs.items():
    for char in invalid_chars:
        department = department.replace(char, '_')
    sheetname = department[:31]  # Sheet names can't be longer than 31 characters or Excel breaks
    wb.create_sheet(title=sheetname)
    ef.create_pie_charts(df, wb, sheetname)

# Save changes to the sheet
wb.save("analysis.xlsx")

# Clean up temp file
try:
    os.remove('combined.xlsx')
except:
    print("WARN: could not delete temp file.")

input("\n\nAll tasks completed. Press enter to close.")
