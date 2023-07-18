import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
from PIL import ImageTk, Image
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from tkinter import filedialog, messagebox


# does not bring in missing items, and the columns brought in for V LOOKup are bringing in slightly different values

def add_active_award_file():
    file_names = ["Active Contract File", "Prev Contract", "Awards", "Backlog", "Sales History", "SND", "VPC",
                  "Running File - 30 Day Notice Contract Price Increase_Sager - COSTED"]
    file_paths = []

    for file_name in file_names:
        file_path = filedialog.askopenfilename(title="Select {} file".format(file_name))
        if not file_path:
            messagebox.showerror("Error", "File selection cancelled.")
            return
        file_paths.append(file_path)

    active_award_file_path = file_paths[0]
    active_award_workbook = load_workbook(active_award_file_path)

    # Add columns to Active Supplier Contracts sheet and add more columns into the sheet we need
    active_sheet = active_award_workbook["Active Supplier Contracts"]
    column_names = ["GP%", "Cost", "Cost Note", "Quote#", "Cost Exp Date", "Cost MOQ", "Prev Contract MPN",
                    "Prev Contract Price",
                    "MPN Match", "Price Match MPN", "LAST WEEK Contract Change", "Contract Change", "PSoft Part",
                    "count",
                    "SUM", "AVG", "DIFF", "PSID All Contract Prices Same?", "PS Award Price", "PS Award Exp Date",
                    "PS Awd Cust ID",
                    "Price Match Award", "Corp Awd Loaded", "Review Note", "90 DAY PI - NEW PRICE", "PI SENT DATE",
                    "DIFF Price Increase",
                    "PI EFF DATE", "12 Month CPN Sales", "DIFF LW", "LW Cost", "LW Cost Note", "LW Cost Exp Date",
                    "LW Review Note", "Estimated $ Value",
                    "Estimated Cost$", "Estimated GP$", "GL-Interconnect Qte - Feb (Y/N)", "DS-Battery Qte - Mar (Y/N)",
                    "Part Class", "Sager Stock",
                    "Cost to Use 1", "Resale 1", "Price Match", "Sager Min", "Min Match", "New Special Cost",
                    "Internal Comments", "New Special Quote#",
                    "SP Exp Date", "Gil Rev Price", "Gil Rev Margin", "Gil Rev MOQ", "Gil Rev SPQ",
                    "Gil Rev Price Match", "Price OK", "Min OK",
                    "BOM COMMENT", "Status", "Assigned"]

    columns_length = len(active_sheet[1])  # Get the length of the first row (columns count)
    for i, column_name in enumerate(column_names, start=1):
        active_sheet.cell(row=1, column=i + columns_length).value = column_name

    # Load active and prev contract dataframes
    active_df = pd.read_excel(active_award_file_path, header=1)
    prev_contract_file_path = file_paths[1]
    prev_df = pd.read_excel(prev_contract_file_path, header=0)

    # Convert GP% to a percentage in the 'Prev Contract' dataframe
    prev_df["GP%"] = prev_df["GP%"].apply(lambda x: x / 100)

    # Create 'Lost Items' sheet
    lost_items_df = prev_df[~prev_df['IPN'].isin(active_df['IPN'])]
    lost_items_sheet = active_award_workbook.create_sheet('Lost Items')
    for r in dataframe_to_rows(lost_items_df, index=False, header=True):
        lost_items_sheet.append(r)

    # Load and merge data from other files
    for file_path, file_name in zip(file_paths[1:], file_names[1:]):  # Skip active_award_file
        data = pd.read_excel(file_path)

        if file_name == "Prev Contract":  # Perform VLOOKUP-like operation for Prev Contract file
            for row in range(2, active_sheet.max_row + 1):
                ipn = active_sheet.cell(row=row, column=1).value
                matching_row = data[data["IPN"] == ipn]
                if not matching_row.empty:
                    for i, column_name in enumerate(column_names, start=1):
                        cell = active_sheet.cell(row=row, column=i + columns_length)
                        cell.value = matching_row[column_name].values[0]

                        # Set the number format for the 'GP%' and 'Cost' columns
                        if column_name == "GP%":
                            cell.number_format = '0.00%'
                        elif column_name == "Cost":
                            cell.number_format = '$#,##0.00'

        # Create a new sheet for each file and add data to it
        new_sheet = active_award_workbook.create_sheet(title=file_name)
        for r in dataframe_to_rows(data, index=False, header=True):
            new_sheet.append(r)
    try:
        active_award_workbook.save(active_award_file_path)
        messagebox.showinfo("Success!", "Files merged, columns added, and VLOOKUP completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))


class ExcelSorter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Sorting Creation Files")
        self.window.configure(bg="white")
        self.window.geometry("1800x750")
        self.file_paths = []
        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.configure("TButton", font=("Times New Roman", 14, "bold"), width=60, height=2)
        style.map("TButton",
                  foreground=[('active', 'red')],
                  background=[('active', 'blue')])

        title_label = ttk.Label(self.window, text="Welcome Partnership Member!",
                                font=("Times New Roman", 30, "underline"), background="white")
        title_label.pack(pady=10)

        logo_image = Image.open('images/electronic.png')
        logo_image = logo_image.resize((200, 200), Image.ANTIALIAS)
        logo_photo = ImageTk.PhotoImage(logo_image)
        logo_label = ttk.Label(self.window, image=logo_photo, background="white")
        logo_label.image = logo_photo
        logo_label.place(x=1575, y=0)

        description_label = ttk.Label(self.window,
                                      text="-This tool allows you to sort your Excel files for our Creation "
                                           "Contact-",
                                      font=("Times New Roman", 14), background="white")
        description_label.pack(pady=10)

        sort_award_button = ttk.Button(self.window, text="Sort Award File", command=self.sort_award_file,
                                       style="TButton")
        sort_award_button.pack(pady=10)

        sort_backlog_button = ttk.Button(self.window, text="Sort Backlog File", command=self.sort_backlog_file,
                                         style="TButton")
        sort_backlog_button.pack(pady=10)

        sort_last_ship_date_button = ttk.Button(self.window, text="Sort Sales History File",
                                                command=self.sort_by_last_ship_date, style="TButton")
        sort_last_ship_date_button.pack(pady=10)

        sort_ship_and_debit = ttk.Button(self.window, text="Sort SND File", command=self.sort_ship_and_debit,
                                         style="TButton")
        sort_ship_and_debit.pack(pady=10)

        sort_vpc = ttk.Button(self.window, text="Sort VPC File", command=self.sort_vpc, style="TButton")
        sort_vpc.pack(pady=10)

        add_instructions_for_active_contracts_file = ttk.Label(self.window,
                                                               text="This last button will allow you to merge your "
                                                                    "files accordingly."
                                                                    " Order to select files is: Current Contract, "
                                                                    "Previous Weeks Contract, Awards, Backlog, "
                                                                    "Sales History, SND, VPC, Running File",
                                                               font=("Times New Roman", 16), background="white")
        add_instructions_for_active_contracts_file.pack(pady=10)

        add_active_award_button = ttk.Button(self.window, text="Prepare Your Active Contracts File",
                                             command=add_active_award_file, style="TButton")
        add_active_award_button.pack(pady=10)

        logo_label = ttk.Label(self.window, background="white")
        logo_label.pack(pady=10)

        logo_image = Image.open('images/Sager-logo.png')
        logo_image = ImageTk.PhotoImage(logo_image)
        logo_label.config(image=logo_image)
        logo_label.image = logo_image

    @staticmethod
    def select_file():
        file_path = filedialog.askopenfilename(title="Select Excel file",
                                               filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")))

        if file_path:
            return file_path
        else:
            return None

    def sort_award_file(self):
        file_path = self.select_file()
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'Award Cust ID'], [True, False])

    def sort_backlog_file(self):
        file_path = self.select_file()
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'Backlog Entry'], [True, False])

    def sort_by_last_ship_date(self):
        file_path = self.select_file()
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'Last Ship Date'], [True, False])

    def sort_ship_and_debit(self):
        file_path = self.select_file()
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'SND Cost'], [True, True])

    def sort_vpc(self):
        file_path = self.select_file()
        if file_path:
            self.sort_excel(file_path, ['PART ID', 'VPC Cost'], [True, False])

    @staticmethod
    def sort_excel(file_path, sort_columns, ascending_order):
        if not sort_columns:
            messagebox.showerror("Error", "No columns selected for sorting.")
            return

        try:
            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(file_path)

            # If 'SND Cost' is one of the sort columns, convert it to numeric
            if 'SND Cost' in sort_columns:
                df['SND Cost'] = pd.to_numeric(df['SND Cost'], errors='coerce')

            # If 'VPC Cost' is one of the sort columns, convert it to numeric
            if 'VPC Cost' in sort_columns:
                df['VPC Cost'] = pd.to_numeric(df['VPC Cost'], errors='coerce')

            # Sort the DataFrame based on the selected columns
            df = df.sort_values(by=sort_columns, ascending=ascending_order)

            # Save the sorted DataFrame back to the Excel file
            df.to_excel(file_path, index=False)

            messagebox.showinfo("Success!", "Excel file sorted and saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    @staticmethod
    def write_data_to_sheet(sheet, df):
        for r in dataframe_to_rows(df, index=False, header=True):
            sheet.append(r)

    def run(self):
        self.window.mainloop()


# Create an instance of the ExcelSorter and run the program
sorter = ExcelSorter()
sorter.run()
