from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import pandas as pd
import re


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

def split_and_explode(df, column, delimiter=';'):
    """
    Splits the values in a column based on a delimiter and returns a DataFrame where each row contains a single value from the split data.

    Args:
        df (DataFrame): The DataFrame containing the column to split.
        column (str): The column to split.
        delimiter (str): The delimiter to use for splitting the column.

    Returns:
        DataFrame: A DataFrame where each row contains a single value from the split data.
    """
    s = df[column].str.split(delimiter).apply(pd.Series, 1).stack()
    s.index = s.index.droplevel(-1)
    s.name = column
    return df.drop(columns=column).join(s)


def get_responsibility_and_member_of_data(df):
    """
    Retrieves responsibility and 'Member of' data from a given DataFrame.

    Args:
        df (DataFrame): A pandas DataFrame to retrieve responsibility and 'Member of' data from.

    Returns:
        DataFrame: A DataFrame containing responsibility and 'Member of' data.

    Raises:
        KeyError: If 'RESPONSIBILITY_NAME' or 'MEMBER_OF' column doesn't exist in the DataFrame.
    """
    responsibilities = df['RESPONSIBILITY_NAME'].value_counts().reset_index()
    responsibilities.columns = ['RESPONSIBILITY_NAME', 'COUNTS']

    member_of_df = split_and_explode(df, 'MEMBER_OF')
    member_of = member_of_df['MEMBER_OF'].value_counts().reset_index()
    member_of.columns = ['MEMBER_OF', 'COUNTS_MEMBER_OF']

    return responsibilities, member_of


def create_pie_charts(df, ws):
    """
    Creates pie charts in a worksheet. It also contains formatting data.

    Args:
        df (DataFrame): A pandas DataFrame containing responsibility and 'Member of' data.
        ws (Worksheet): An openpyxl worksheet object.
    """
    responsibilities, member_of = get_responsibility_and_member_of_data(df)
    
    append_dataframe_to_sheet(ws, responsibilities)
    adjust_column_width(ws, {'A': 45, 'B': 10})
    align_cells(ws, ['B'], Alignment(horizontal='center'))
    create_pie_chart(ws, responsibilities)

    # Add some space between the two tables
    ws.append([])
    
    append_dataframe_to_sheet(ws, member_of)
    adjust_column_width(ws, {'A': 45, 'B': 10})
    align_cells(ws, ['B'], Alignment(horizontal='center'))
    create_pie_chart(ws, member_of)



def sanitize_sheet_name(sheet_name):
    invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '')
    sheet_name = sheet_name.replace(' ', '_')
    sheet_name = sheet_name[:31]  # Trim to 31 characters
    return sheet_name


def create_job_title_sheets_and_charts(df, wb):
    """
    Creates a worksheet for each job title in the DataFrame.
    Each worksheet contains a pie chart and a table with responsibility data.

    Args:
        df (DataFrame): A pandas DataFrame containing user responsibility data.
        wb (Workbook): An openpyxl workbook object.
    """
    for job_title, group in df.groupby('JOB_TITLE'):
        sanitized_job_title = sanitize_sheet_name(job_title)
        ws = wb.create_sheet(title=sanitized_job_title)  # Save the reference to the created sheet
        create_pie_charts(group, ws)  # Pass the sheet directly

