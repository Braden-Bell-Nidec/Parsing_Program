# README - Outlier Detection in Company Role Distribution 

## Overview

This Python script is designed to analyze role distribution within a company by identifying potential outliers based on percentage thresholds. The user provides an Excel file (`.xlsx`) with Enterprise Project Governance Architecture (EPGA) data and a Comma Separated Value (CSV) file with Active Directory (AD) data. The script merges these files, reads the data, and allows the user to specify a percentage threshold. The script then identifies outliers, or roles that constitute less than the specified threshold of a department's distribution.

## Prerequisites 

You need to have the following software installed in order to run the script:

1. Python 3.x 
2. The following Python libraries: pandas, openpyxl, and the custom modules Excel_Functions and Combiner (make sure these files are in the same directory as the main script).

## How to run the script

1. Open a terminal or command prompt.
2. Navigate to the directory where the script is located.
3. Run the script by typing `python script_name.py` where "script_name" is the name of the Python script.
4. You will be prompted to provide the paths to the EPGA and AD files. Ensure that the files are accessible and the paths are correctly typed.
5. The script will ask for an outlier percentage threshold. This is the cutoff below which roles are considered outliers. If no value is entered, the default value of 7% will be used.
6. The script will then execute the analysis and present the outliers and non-outliers in the terminal. 
7. Finally, the script generates an Excel file (`analysis.xlsx`) which contains the analysis results with the outliers and non-outliers data in separate sheets. The workbook also includes separate sheets and charts for each unique job title.

## Error Handling 

The script includes comprehensive error handling to guide users when issues arise. Errors may occur due to incorrect file paths, permission errors (if files are open in another program or inaccessible), and data-related issues.

## Output

The script produces an Excel file (`analysis.xlsx`) with the analysis results:

- An 'Outliers' sheet that lists all users who fall into outlier categories (job roles constituting less than the specified percentage of a department's distribution).
- A 'Non-Outliers' sheet that lists all users who don't fall into outlier categories.
- Separate sheets and charts for each unique job title.

## Clean Up 

The script creates a temporary combined Excel file during the execution, which is deleted upon completion.

## Credits
This script was created by Braden Bell with assistance from chatGPT-4