from tkinter import filedialog, Tk, StringVar, IntVar, BooleanVar, Text, Scrollbar, END, DISABLED
from tkinter.ttk import Frame, Label, Entry, Checkbutton, Button


class Application(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.grid(sticky="nsew")  # Use grid instead of pack
        self.master.columnconfigure(0, weight=1)  # Allow column to expand
        self.master.rowconfigure(0, weight=1)  # Allow row to expand
        self.columnconfigure(0, weight=1)  # Allow column to expand
        self.rowconfigure(1, weight=1)  # Allow row with console to expand
        self.create_widgets()

    def create_widgets(self):
        # EPGA
        self.EPGA_label = Label(self, text="EPGA File: ")
        self.EPGA_label.grid(row=0, column=0, sticky="e")
        self.EPGA_var = StringVar()
        self.EPGA_entry = Entry(self, textvariable=self.EPGA_var)
        self.EPGA_entry.grid(row=0, column=1, sticky="ew")
        self.EPGA_browse = Button(self, text="Browse", command=self.browse_EPGA)
        self.EPGA_browse.grid(row=0, column=2)

        # Active Directory
        self.AD_label = Label(self, text="Active Directory File: ")
        self.AD_label.grid(row=1, column=0, sticky="e")
        self.AD_var = StringVar()
        self.AD_entry = Entry(self, textvariable=self.AD_var)
        self.AD_entry.grid(row=1, column=1, sticky="ew")
        self.AD_browse = Button(self, text="Browse", command=self.browse_AD)
        self.AD_browse.grid(row=1, column=2)

        # Checkbox
        self.delete_var = BooleanVar(value=True)
        self.delete_check = Checkbutton(self, text="Delete combined.xlsx after use", variable=self.delete_var)
        self.delete_check.grid(row=2, column=0, columnspan=3)

        # User percentage
        self.user_percent_label = Label(self, text="Custom percentage: ")
        self.user_percent_label.grid(row=3, column=0, sticky="e")
        self.user_percent_var = StringVar()
        self.user_percent_entry = Entry(self, textvariable=self.user_percent_var)
        self.user_percent_entry.grid(row=3, column=1, sticky="ew")

        # Console output
        self.console = Text(self, state=DISABLED)
        self.console.grid(row=4, column=0, columnspan=3, sticky="nsew")

        # Button
        self.run_button = Button(self, text="RUN", command=self.run_parser)
        self.run_button.grid(row=5, column=0, columnspan=3)

    def browse_EPGA(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])  # Only allow .xlsx files
        self.EPGA_var.set(filename)

    def browse_AD(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])  # Only allow .csv files
        self.AD_var.set(filename)

    def run_parser(self):
        # Your code here
        pass


def start():
    root = Tk()
    root.title("Parser Application")
    app = Application(master=root)
    app.mainloop()
    return app


app = start()
