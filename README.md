# README.md

## Python Program: Data Analysis Tool

This is a Python application that merges two Excel files, performs data analysis and produces an output Excel workbook that presents this data along with charts.

### Dependencies

The program uses the following Python libraries:

- `pandas`
- `openpyxl`
- `tkinter`
- `time`
- `sys`

It also uses custom modules named `Excel_Functions`, `GUI` and `Combiner`.

### Usage

The main function of the program is `main(EPGA_File, AD_File, user_percent, delete_combined, progress, status)`, which performs all the core operations. It uses a graphical user interface (GUI) for input.

#### Parameters:

- `EPGA_File`: This is the path of the first Excel file.
- `AD_File`: This is the path of the second Excel file.
- `user_percent`: This is the user-defined percentage threshold for outliers. If left empty, the default threshold of 7% is used. If the input percentage is over 100, the program reverts to the default.
- `delete_combined`: This is a boolean flag to indicate whether to delete the temporary combined file created during processing.
- `progress`: This is a GUI progress bar.
- `status`: This is a GUI status label that provides updates to the user.

### Output

The output is an Excel file named 'analysis.xlsx'. It contains several sheets:

- `Outliers`: This sheet lists users who have been identified as outliers according to the user-defined or default threshold.
- `Non-Outliers`: This sheet lists data that do not meet the outlier criteria.
- Individual sheets for each unique job title, which include charts for further visualization.

### Exception Handling

The program includes comprehensive exception handling for errors such as permission issues, Unicode decoding errors, file not found errors, and general exceptions. The user is provided with descriptive error messages to diagnose problems and is advised to close the program when encountering these exceptions.

### GUI Interface

The application uses a tkinter GUI for input. The `GUI(root, main)` function runs the GUI.

### Installation

To use this program:

1. Ensure that Python 3.6+ and all the required libraries are installed.
2. Download or clone this repository to your local machine.
3. Run the program in a Python environment or using a Python IDE.

Make sure you have the necessary input Excel files and that they are in the correct format.

This program was last updated on May 31, 2023.

### Disclaimer

Please note that the program deletes the temporary combined Excel file created during processing, depending on the `delete_combined` argument. Please be aware of this, as it may lead to data loss if not handled properly. Additionally, the output file 'analysis.xlsx' overwrites any previous file with the same name in the working directory. Please save any important files with a different name or in a different directory to avoid unintentional data loss.
