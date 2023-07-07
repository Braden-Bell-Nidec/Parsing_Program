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
    ad['SAM Account Name'] = ad.apply(adjust_username, axis=1)
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
    #print(combined)
    # Write the combined DataFrame to a new Excel file
    combined.to_excel(output_file, index=False)
    return output_file



def adjust_username(row):
    """
    Adjusts the 'SAM Account Name' for a row of data from the Active Directory file.

    For users from the "EPG Reynosa" office, a username is generated based on the 'Display Name'.
    For other users, the 'SAM Account Name' is simply converted to uppercase.

    Args:
        row (pd.Series): A row of data from the Active Directory file.

    Returns:
        str: The adjusted username.
    """
    if row['Office'] == "EPG Reynosa":
        try:
            # Initialize an empty string for the sanitized name
            sanitized_name = ""

            # Iterate through each character in the name
            for e in row['Display Name']:
                # If the character is a letter, a space, or a comma, add it to the sanitized name
                if e.isalpha() or e.isspace() or e == ',':
                    sanitized_name += e

            # Split the sanitized name into parts
            name_parts = sanitized_name.split(',')
            last_name = name_parts[0].strip().split(" ")[0]
            first_name = name_parts[1].strip().split(" ")[0]
            # Construct username based on the first 6 characters of the last name and the first 2 characters of the first name, converted to uppercase
            username = (last_name[:6] + first_name[:2]).upper()
            return username
        except Exception as e:
            print(f"Error processing name {row['Display Name']}: {e}")
            # If an exception occurred while generating the username, return the original 'SAM Account Name' converted to uppercase
            return row['SAM Account Name'].upper()
    else:
        # If the user is not from the "EPG Reynosa" office, return the 'SAM Account Name' converted to uppercase
        return row['SAM Account Name'].upper()




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

