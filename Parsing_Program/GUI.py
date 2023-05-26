# GUI.py
import tkinter as tk
from tkinter import filedialog, messagebox

def select_files():
    EPGA_File = ''
    AD_File = ''
    root = tk.Tk()
    delete_file = tk.IntVar()  # checkbox variable (0 - not checked, 1 - checked)
    userPercent = tk.StringVar()  # entry field variable
    EPGA_path = tk.StringVar()  # EPGA file path label variable
    AD_path = tk.StringVar()  # AD file path label variable

    def run_script():
        nonlocal EPGA_File
        nonlocal AD_File
        if EPGA_File and AD_File:
            root.destroy()  # close GUI
            return EPGA_File, AD_File, bool(delete_file.get()), userPercent.get()
        else:
            messagebox.showinfo("Error", "Please select both files.")

    def select_EPGA_File():
        nonlocal EPGA_File
        EPGA_File = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        EPGA_path.set(EPGA_File)  # update EPGA file path label

    def select_AD_File():
        nonlocal AD_File
        AD_File = filedialog.askopenfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        AD_path.set(AD_File)  # update AD file path label

    root.title("File Selector")

    tk.Label(root, text="EPGA File:").grid(row=0, column=0, sticky='e')
    tk.Label(root, text="Active Directory File:").grid(row=1, column=0, sticky='e')
    tk.Label(root, text="Outlier Percentage Threshold (default is 7%):").grid(row=2, column=0, sticky='e')
    EPGA_label = tk.Label(root, textvariable=EPGA_path)  # bind EPGA_path variable
    EPGA_label.grid(row=0, column=1)
    AD_label = tk.Label(root, textvariable=AD_path)  # bind AD_path variable
    AD_label.grid(row=1, column=1)
    userPercent_entry = tk.Entry(root, textvariable=userPercent)
    userPercent_entry.grid(row=2, column=1)
    tk.Button(root, text="Browse", command=select_EPGA_File).grid(row=0, column=2)
    tk.Button(root, text="Browse", command=select_AD_File).grid(row=1, column=2)
    tk.Checkbutton(root, text="Delete combined.xlsx", variable=delete_file).grid(row=3, columnspan=2)
    tk.Button(root, text="Run script", command=run_script).grid(row=4, columnspan=3)

    root.mainloop()

    return EPGA_File, AD_File, bool(delete_file.get()), userPercent.get()
