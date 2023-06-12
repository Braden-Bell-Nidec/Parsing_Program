import pandas as pd
import os
def merge_files(epga_file, ad_file, output_file):
    """
    Merge the provided EPGA Excel file and Active Directory CSV file into a single output file.

    Args:
        epga_file (str): The name of the EPGA Excel file.
        ad_file (str): The name of the Active Directory CSV file.
        output_file (str): The name of the output file where the merged data will be written.

    Returns:
        str: The name of the output file containing the merged data.

    Raises:
        ValueError: If either the EPGA file or the Active Directory file is not formatted correctly.
    """

    try:
        # Read the EPGA Excel file
        epga = pd.read_excel(epga_file)
    except pd.errors.ParserError:
        raise ValueError(f"The EPGA file '{epga_file}' is not formatted correctly.\nPlease ensure it is a valid Excel file with expected column names.")

    try:
        # Read the Active Directory CSV file
        ad = pd.read_csv(ad_file)
    except pd.errors.ParserError:
        raise ValueError(f"The Active Directory file '{ad_file}' is not formatted correctly.\nPlease ensure it is a valid CSV file with expected column names.")

    # Convert the "SAM Account Name", 'Member of', and 'Office' columns to uppercase for case-insensitive match
    ad['SAM Account Name'] = ad['SAM Account Name'].str.upper()
    ad['Member of'] = ad['Member of'].str.upper()
    ad['Office'] = ad['Office'].str.upper()

    # Merge the two DataFrames on the user name columns
    combined = pd.merge(epga, ad, left_on='USER_NAME', right_on='SAM Account Name', how='inner')

    # Select the columns of interest
    combined = combined[['Department', 'USER_NAME', 'RESPONSIBILITY_NAME', 'Title', 'Member of', 'Office']]

    # Replace '-' values with "Unknown"
    combined['Department'] = combined['Department'].replace('-', 'Unknown')
    combined['Title'] = combined['Title'].replace('-', 'Unknown')
    combined['Office'] = combined['Office'].replace('-', 'Unknown')

    # Rename the columns to the desired names
    combined.columns = ['DEPARTMENT', 'USER_NAME', 'RESPONSIBILITY_NAME', 'JOB_TITLE', 'MEMBER_OF', 'OFFICE']

    # Write the combined DataFrame to a new Excel file
    combined.to_excel(output_file, index=False)
    
    return output_file



import os

def delete_temp(tempFile):
    """
    Attempts to delete the specified temporary file.

    Args:
        tempFile (str): The name of the temporary file to delete.

    Returns:
        None.

    Prints:
        A warning message if the temporary file could not be deleted.
    """

    try:
        # Attempt to remove the temporary file
        os.remove(tempFile)
    except:
        # If an exception occurs during file deletion, print a warning message
        print("WARN: Could not delete temp file.")

