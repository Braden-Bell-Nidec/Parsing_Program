import pandas as pd

def merge_files(epga_file, ad_file, output_file):
    # Load both csv files into pandas DataFrames
    df_epga = pd.read_excel(epga_file)
    df_ad = pd.read_csv(ad_file)

    # Convert the "SAM Account Name" column to uppercase for case-insensitive match
    df_ad['SAM Account Name'] = df_ad['SAM Account Name'].str.upper()

    # Merge the two DataFrames on the user name columns
    df_combined = pd.merge(df_epga, df_ad, left_on='USER_NAME', right_on='SAM Account Name', how='inner')

    # Select the columns we are interested in
    df_combined = df_combined[['Department', 'USER_NAME', 'RESPONSIBILITY_NAME', 'Title']]

    # Rename the columns to the desired names
    df_combined.columns = ['DEPARTMENT', 'USER_NAME', 'RESPONSIBILITY_NAME', 'JOB_TITLE']

    # Write the combined DataFrame to a new Excel file
    df_combined.to_excel(output_file, index=False)

    return output_file
