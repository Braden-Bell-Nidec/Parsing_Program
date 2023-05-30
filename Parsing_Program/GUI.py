import tkinter as tk
from tkinter import filedialog, Text, IntVar, StringVar
import sys
import threading

class TextboxWriter:
    def __init__(self, textbox):
        self.textbox = textbox
        self.textbox.configure(state='disabled')

    def write(self, text):
        self.textbox.configure(state='normal')
        self.textbox.insert(tk.END, text)
        self.textbox.configure(state='disabled')


    def flush(self):
        pass


class GUI:
    def __init__(self, root, main_func):
        self.root = root
        self.main_func = main_func

        self.root.title("Parser by Braden Bell")

        # Create a frame that fills the root window
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill="both", expand=True)

        # File selectors
        self.EPGA_file = StringVar()
        self.AD_file = StringVar()
        self.create_file_selector(self.frame, "Select EPGA file", self.EPGA_file, 0)
        self.create_file_selector(self.frame, "Select Active Directory file", self.AD_file, 1)

        # Input box for user percentage
        self.user_percentage = StringVar()
        self.create_input_box(self.frame, "Enter outlier percentage threshold (default is 7%): ", self.user_percentage, 2)

        # Checkbox
        self.delete_combined = IntVar(value=1)
        self.create_checkbox(self.frame, "Delete combined.xlsx upon completion", self.delete_combined, 3)

        # Run button
        self.create_run_button(self.frame, "Run", self.run_program, 4)

        # Text box for program output
        self.output_box = self.create_text_box(self.frame, 5)
        sys.stdout = TextboxWriter(self.output_box)

    def create_file_selector(self, parent, label_text, string_var, row):
        frame = tk.Frame(parent)
        label = tk.Label(frame, text=label_text)
        entry = tk.Entry(frame, textvariable=string_var)
        button = tk.Button(frame, text="Browse", command=lambda: string_var.set(filedialog.askopenfilename(filetypes=(("xlsx files", "*.xlsx"), ("csv files", "*.csv"),("all files", "*.*")))))
        label.pack(side="left")
        entry.pack(side="left")
        button.pack(side="left")
        frame.grid(row=row, column=0, sticky="ew")
        parent.grid_rowconfigure(row, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    def create_input_box(self, parent, label_text, string_var, row):
        frame = tk.Frame(parent)
        label = tk.Label(frame, text=label_text)
        entry = tk.Entry(frame, textvariable=string_var)
        label.pack(side="left")
        entry.pack(side="left")
        frame.grid(row=row, column=0, sticky="ew")
        parent.grid_rowconfigure(row, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    def create_checkbox(self, parent, label_text, int_var, row):
        checkbox = tk.Checkbutton(parent, text=label_text, variable=int_var)
        checkbox.grid(row=row, column=0, sticky="w")
        parent.grid_rowconfigure(row, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    def create_run_button(self, parent, label_text, command, row):
        button = tk.Button(parent, text=label_text, command=command)
        button.grid(row=row, column=0, sticky="ew")
        parent.grid_rowconfigure(row, weight=1)
        parent.grid_columnconfigure(0, weight=1)

    def create_text_box(self, parent, row):
        text_box = Text(parent)
        text_box.grid(row=row, column=0, sticky="nsew")
        parent.grid_rowconfigure(row, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        return text_box


    def run_program(self):
        EPGA_File = self.EPGA_file.get()
        AD_File = self.AD_file.get()
        user_percent = self.user_percentage.get()
        delete_temp = self.delete_combined.get()
        threading.Thread(target=self.main_func, args=(EPGA_File, AD_File, user_percent, delete_temp)).start()
