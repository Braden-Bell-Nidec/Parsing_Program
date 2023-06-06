from openpyxl.chart import PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import pandas as pd

def sanitize_sheet_name(sheet_name):
    invalid_chars = ['_', '-', ',', '\\', '/', '*', '[', ']', ':', '?', ' ']
    abbreviation_dict = {
        'Manager': 'Mgr',
        'manager': 'Mgr',
        'Manger': 'Mgr',
        'Associate': 'Assoc',
        'I': '1',
        'II': '2',
        'III': '3',
        'InformationTechnology': 'IT',
        'Technician': 'Tech',
        'Mechanical' : 'Mech',
        'Certification' : 'Cert',
        'Senior' : 'Sr.',
        'HumanResources': 'HR',
        'HumanResouce': 'HR',
        'BuisinessPartner': 'BP',
        'President': 'Pres',
        'Engineer': 'Eng',
        'Engineering': 'Eng',
        'Operations': 'Op',
        'MANKATO': 'MKTO',
        'LEXINGTON': 'LEX',
        'REYNOSA': 'REY',
        'Electronic': 'Elec',
        'Component': 'Comp',
        'Assembler': 'Assem',
        'Marketing': 'MkTg',
        'Maintenance': 'Maint.',
        'And': '&',
        'and': '&',
        'General': 'Gen',
        'Network': 'Net',
        'Infrastructure': 'Inf',
        'Specialist': 'Spclst',
        'Environmental': 'Enviro',
        'Health': 'Hlth',
        'Safety': 'Sfty',
        'Director': 'Dir',
        'Administrative': 'Admin',
        'Administrator': 'Admin',
        'Product': 'Prod',
        'Accounts': 'Accts',
        'Aftermarket': 'AM',
        'Remanufacturing': 'Reman',
        'Supervisor': 'Suprvsr.',
        'Development': 'Dev',
        'Caterpillar': 'CAT',
        'Communications': 'Comms',
        'Logistics': 'Log',
        'Compliance': 'Cmpl',
        'Shipping': 'Ship',
        'Recieving': 'Rec',
        'Technical': 'Tech',
        'Fabrication': 'Fab',
        'Manufacturing': 'Man.',
        'Represenative': 'Rep'
        }

    for word in sheet_name.split():
        if word in abbreviation_dict:
            sheet_name = sheet_name.replace(word, abbreviation_dict[word])

    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '')
    
    #If sheet_name is empty string after removing invalid chars, replace it with "Unnamed"
    if sheet_name == "":
        sheet_name = "Unnamed"
    
    #Truncate to 31 characters
    sheet_name = sheet_name[:31]

    return sheet_name

def append_dataframe_to_sheet(ws, df, start_row=1, start_col=1):
    """
    Appends a pandas DataFrame to a worksheet. DataFrame values start at the specified row and column.

    Args:
        ws (Worksheet): An openpyxl worksheet object.
        df (DataFrame): A pandas DataFrame.
        start_row (int): The row index where the dataframe should start.
        start_col (int): The column index where the dataframe should start.
    """

    #Write data from the DataFrame to the worksheet
    for i, row in enumerate(df.values, start=1):  #Starting index from 1
        for j, item in enumerate(row, start=1):  #Starting index from 1
            ws.cell(row=start_row+i-1, column=start_col+j-1, value=item)

    #Write column names to the worksheet
    for j, col_name in enumerate(df.columns, start=1):  #Starting index from 1
        ws.cell(row=start_row, column=start_col+j-1, value=col_name)


def create_job_title_sheets_and_charts(df, wb):
    job_titles = df['JOB_TITLE'].unique()
    for title in job_titles:
        job_title_df = df[df['JOB_TITLE'] == title]
        offices = job_title_df['OFFICE'].unique()
        for office in offices:
            sanitized_title = sanitize_sheet_name(title)
            sanitized_office = sanitize_sheet_name(office)
            combined_title = (sanitized_office + "-" + sanitized_title)[:31] 
            ws = wb.create_sheet(title=combined_title)
            office_group = job_title_df[job_title_df['OFFICE'] == office]
            create_pie_charts(office_group, ws)



    #If the default 'Sheet' exists, remove it
    if 'Sheet' in wb:
        wb.remove(wb['Sheet'])


def adjust_column_width(sheet, cols_width_dict):
    """
    Adjusts the column width of specified columns in a worksheet.

    Args:
        sheet (Worksheet): An openpyxl worksheet object.
        cols_width_dict (dict): A dictionary containing column letters as keys and widths as values.
    """
    #Iterate over each column 
    for col, width in cols_width_dict.items():
        #And set width
        sheet.column_dimensions[col].width = width

def align_cells(sheet, cols, alignment):
    """
    Aligns cells in specified columns to a given alignment.

    Args:
        sheet (Worksheet): An openpyxl worksheet object.
        cols (list): A list of column letters to be aligned.
        alignment (Alignment): An openpyxl Alignment object.
    """
    #Iterate over each column
    for col in cols:
        #iterate over each cell in current column
        for cell in sheet[col]:
            #set alignment of current cell
            cell.alignment = alignment


def create_excel_pie_chart(sheet, df, min_col, max_col, chart_location):
    chart = PieChart()
    labels = Reference(sheet, min_col=min_col, min_row=2, max_row=len(df)+1) # Adjusted min_row and max_row
    data = Reference(sheet, min_col=max_col, min_row=2, max_row=len(df)+1) # Adjusted min_row and max_row
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = 'Responsibility Distribution' if min_col==1 else 'Member Of Distribution'
    sheet.add_chart(chart, chart_location)




def create_pie_charts(df, ws):
    responsibilities, member_of = get_responsibility_and_member_of_data(df)
    append_dataframe_to_sheet(ws, responsibilities, start_row=1, start_col=1)
    adjust_column_width(ws, {'A': 45, 'B': 10})
    align_cells(ws, ['B'], Alignment(horizontal='center'))
    create_excel_pie_chart(ws, responsibilities, min_col=1, max_col=2, chart_location="F1")

    append_dataframe_to_sheet(ws, member_of, start_row=1, start_col=4)
    adjust_column_width(ws, {'D': 45, 'E': 20})
    align_cells(ws, ['E'], Alignment(horizontal='center'))
    create_excel_pie_chart(ws, member_of, min_col=4, max_col=5, chart_location="F16")



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








