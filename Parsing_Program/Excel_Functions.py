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
