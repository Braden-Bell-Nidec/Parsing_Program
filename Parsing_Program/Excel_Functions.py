from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment


def adjust_column_width(sheet, cols_width_dict):
    """
    Adjusts the column width of specified columns in a worksheet.

    Args:
        sheet (Worksheet): An openpyxl worksheet object.
        cols_width_dict (dict): A dictionary containing column letters as keys and widths as values.
    """
    for col, width in cols_width_dict.items():
        sheet.column_dimensions[col].width = width


def align_cells(sheet, cols, alignment):
    """
    Aligns cells in specified columns to a given alignment.

    Args:
        sheet (Worksheet): An openpyxl worksheet object.
        cols (list): A list of column letters to be aligned.
        alignment (Alignment): An openpyxl Alignment object.
    """
    for col in cols:
        for cell in sheet[col]:
            cell.alignment = alignment


def create_pie_chart(sheet, responsibilities):
    """
    Creates a pie chart in the specified worksheet based on the responsibility data.

    Args:
        sheet (Worksheet): An openpyxl worksheet object.
        responsibilities (DataFrame): A pandas DataFrame containing responsibility data.
    """
    chart = PieChart()
    labels = Reference(sheet, min_col=1, min_row=2, max_row=len(responsibilities)+1)
    data = Reference(sheet, min_col=2, min_row=1, max_row=len(responsibilities)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = 'Responsibility Distribution'
    sheet.add_chart(chart, "C1")

def append_dataframe_to_sheet(sheet, df):
    """
    Appends all rows of a DataFrame to a given worksheet.

    Args:
        sheet (Worksheet): An openpyxl worksheet object.
        df (DataFrame): A pandas DataFrame to append.
    """
    for row in dataframe_to_rows(df, index=False, header=True):
        sheet.append(row)


def get_responsibility_data(df):
    """
    Retrieves responsibility data from a given DataFrame.

    Args:
        df (DataFrame): A pandas DataFrame to retrieve responsibility data from.

    Returns:
        DataFrame: A DataFrame containing responsibility data.

    Raises:
        KeyError: If 'RESPONSIBILITY_NAME' column doesn't exist in the DataFrame.
    """
    responsibilities = df['RESPONSIBILITY_NAME'].value_counts().reset_index()
    responsibilities.columns = ['RESPONSIBILITY_NAME', 'COUNTS']
    return responsibilities


def create_pie_charts(df, wb, sheetname):
    """
    Creates pie charts in a worksheet. It also contains formatting data.

    Args:
        df (DataFrame): A pandas DataFrame containing responsibility data.
        wb (Workbook): An openpyxl workbook object.
        sheetname (str): The name of the worksheet to create the pie charts in.
    """
    ws = wb[sheetname]
    responsibilities = get_responsibility_data(df)
    append_dataframe_to_sheet(ws, responsibilities)
    adjust_column_width(ws, {'A': 45, 'B': 10})
    align_cells(ws, ['B'], Alignment(horizontal='center'))
    create_pie_chart(ws, responsibilities)
