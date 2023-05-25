from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import pandas as pd


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


def create_excel_pie_chart(sheet, df, min_col, max_col, chart_location):
    """
    Creates a pie chart in the specified worksheet based on the responsibility data.

    Args:
        sheet (Worksheet): An openpyxl worksheet object.
        df (DataFrame): A pandas DataFrame containing responsibility data.
        min_col (int): Minimum column index for the data to be charted.
        max_col (int): Maximum column index for the data to be charted.
        chart_location (str): Excel-style cell reference for the location of the chart.
    """
    chart = PieChart()
    labels = Reference(sheet, min_col=min_col, min_row=2, max_row=len(df)+1)
    data = Reference(sheet, min_col=max_col, min_row=1, max_row=len(df)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = 'Responsibility Distribution' if min_col==1 else 'Member Of Distribution'
    sheet.add_chart(chart, chart_location)


def append_dataframe_to_sheet(ws, df, start_row=1, start_col=1):
    """
    Appends a pandas DataFrame to a worksheet. DataFrame values start at the specified row and column.

    Args:
        ws (Worksheet): An openpyxl worksheet object.
        df (DataFrame): A pandas DataFrame.
        start_row (int): The row index where the dataframe should start.
        start_col (int): The column index where the dataframe should start.
    """

    # Write data from the DataFrame to the worksheet
    for i, row in enumerate(df.values, start=1):  # Starting index from 1
        for j, item in enumerate(row, start=1):  # Starting index from 1
            ws.cell(row=start_row+i-1, column=start_col+j-1, value=item)

    # Write column names to the worksheet
    for j, col_name in enumerate(df.columns, start=1):  # Starting index from 1
        ws.cell(row=start_row, column=start_col+j-1, value=col_name)


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

    append_dataframe_to_sheet(ws, responsibilities, start_row=2, start_col=2)
    adjust_column_width(ws, {'A': 45, 'B': 10})
    align_cells(ws, ['B'], Alignment(horizontal='center'))
    create_excel_pie_chart(ws, responsibilities, min_col=1, max_col=2, chart_location="F1")

    append_dataframe_to_sheet(ws, member_of, start_row=2, start_col=4)
    adjust_column_width(ws, {'C': 45, 'D': 10})
    align_cells(ws, ['D'], Alignment(horizontal='center'))
    create_excel_pie_chart(ws, member_of, min_col=3, max_col=4, chart_location="F16")


def sanitize_sheet_name(sheet_name):
    invalid_chars = ['\\', '/', '*', '[', ']', ':', '?', ' ']
    abbreviation_dict = {
        'Manager' : 'Mgr',
        'Associate': 'Assoc',
        'I': '1',
        'II': '2',
        'III': '3',
        'Information Technology': 'IT',
        'Technician': 'Tech',
        'Mechanical' : 'Mech',
        'Certification' : 'Cert',
        'Senior' : 'Sr.',
        'Human Resources': 'HR',
        'President': 'Pres',
        'Engineer': 'Eng',
        'Operations': 'Op'
        }

    for word in sheet_name.split():
        if word in abbreviation_dict:
            sheet_name = sheet_name.replace(word, abbreviation_dict[word])

    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '')
    
    # If sheet_name is empty string after removing invalid chars, replace it with "Unnamed"
    if sheet_name == "":
        sheet_name = "Unnamed"
    
    # Truncate to 31 characters
    sheet_name = sheet_name[:31]

    return sheet_name


def create_job_title_sheets_and_charts(df, wb):
    job_titles = df['JOB_TITLE'].unique()

    for title in job_titles:
        #print(f"Original title: {title}")  # Debug line
        sanitized_title = sanitize_sheet_name(title)
        #print(f"Sanitized title: {sanitized_title}")  # Debug line
        ws = wb.create_sheet(title=sanitized_title)

        group = df[df['JOB_TITLE'] == title]

        create_pie_charts(group, ws)

    # If the default 'Sheet' exists, remove it
    if 'Sheet' in wb:
        wb.remove(wb['Sheet'])







