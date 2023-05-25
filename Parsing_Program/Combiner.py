import pandas as pd
import os
def merge_files(epga_file, ad_file, output_file):
    """
    This function merges the provided EPGA Excel file and the Active Directory CSV file.
    
    Args:
        epga_file (str): The name of the EPGA Excel file.
        ad_file (str): The name of the Active Directory CSV file.
        output_file (str): The name of the output file where the merged data will be written.
        
    Returns:
        output_file (str): The name of the output file containing the merged data.
        
    Raises:
        ValueError: If either the EPGA file or the Active Directory file is not formatted correctly.
    """
    try:
        epga = pd.read_excel(epga_file)
    except pd.errors.ParserError:
        raise ValueError(f"The EPGA file '{epga_file}' is not formatted correctly.\nPlease ensure it is a valid Excel file with expected column names.")
    
    try:
        ad = pd.read_csv(ad_file)
    except pd.errors.ParserError:
        raise ValueError(f"The Active Directory file '{ad_file}' is not formatted correctly.\nPlease ensure it is a valid CSV file with expected column names.")

    # Convert the "SAM Account Name" and 'Member of' columns to uppercase for case-insensitive match
    ad['SAM Account Name'] = ad['SAM Account Name'].str.upper()
    ad['Member of'] = ad['Member of'].str.upper()

    # Merge the two DataFrames on the user name columns
    combined = pd.merge(epga, ad, left_on='USER_NAME', right_on='SAM Account Name', how='inner')

    # Select the columns we are interested in
    combined = combined[['Department', 'USER_NAME', 'RESPONSIBILITY_NAME', 'Title', 'Member of']]

    # Replace '-' values with "Unknown"
    combined['Department'] = combined['Department'].replace('-', 'Unknown')
    combined['Title'] = combined['Title'].replace('-', 'Unknown')

    # Rename the columns to the desired names
    combined.columns = ['DEPARTMENT', 'USER_NAME', 'RESPONSIBILITY_NAME', 'JOB_TITLE', 'MEMBER_OF']

    # Write the combined DataFrame to a new Excel file
    combined.to_excel(output_file, index=False)
    return output_file


def delete_temp(tempFile):
    """
    This function attempts to delete the specified temporary file.
    
    Args:
        tempFile (str): The name of the temporary file to delete.
        
    Returns:
        None.
        
    Prints:
        A warning message if the temporary file could not be deleted.
    """
    try:
        os.remove(tempFile)
    except:
        print("WARN: could not delete temp file.")
