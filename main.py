import tkinter as tk
import tkinter.ttk as ttk

import numpy as np
import pandas as pd
from PIL import ImageTk, Image
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from tkinter import filedialog, messagebox


class ExcelSorter:
    def __init__(self):
        self.filename = None
        self.window = tk.Tk()
        self.window.title("Sorting Creation Files For Contract")
        self.window.configure(bg="white")
        self.window.geometry("940x600")

        # Create a canvas and a vertical scrollbar
        self.canvas = tk.Canvas(self.window)
        self.scrollbar = ttk.Scrollbar(self.window, orient="vertical", command=self.canvas.yview)

        # Configure the canvas to respond to the scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a frame to hold your widgets, and add it to the canvas
        self.inner_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((self.window.winfo_width() / 2, 0), window=self.inner_frame, anchor="n")

        # Configure the canvas's scroll-region to encompass the frame
        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Pack the scrollbar, making sure it sticks to the right side
        self.scrollbar.pack(side="right", fill="y")

        # Configure the canvas to expand and fill the window
        self.canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)

        # Canvas - Scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.canvas.yview)

        self.file_paths = []
        self.column_names = ["GP%", "Cost", "Cost Note", "Quote#", "Cost Exp Date", "Cost MOQ", "Prev Contract MPN",
                             "Prev Contract Price",
                             "MPN Match", "Price Match MPN", "LAST WEEK Contract Change", "Contract Change",
                             "PSoft Part",
                             "count",
                             "SUM", "AVG", "DIFF", "PSID All Contract Prices Same?", "PS Award Price",
                             "PS Award Exp Date",
                             "PS Awd Cust ID",
                             "Price Match Award", "Corp Awd Loaded", "Review Note", "90 DAY PI - NEW PRICE",
                             "PI SENT DATE",
                             "DIFF Price Increase",
                             "PI EFF DATE", "12 Month CPN Sales", "DIFF LW", "LW Cost", "LW Cost Note",
                             "LW Cost Exp Date",
                             "LW Review Note", "Estimated $ Value",
                             "Estimated Cost$", "Estimated GP$", "GL-Interconnect Qte - Feb (Y/N)",
                             "DS-Battery Qte - Mar (Y/N)",
                             "Part Class", "Sager Stock",
                             "Cost to Use 1", "Resale 1", "Price Match", "Sager Min", "Min Match", "New Special Cost",
                             "Internal Comments", "New Special Quote#",
                             "SP Exp Date", "Gil Rev Price", "Gil Rev Margin", "Gil Rev MOQ", "Gil Rev SPQ",
                             "Gil Rev Price Match", "Price OK", "Min OK",
                             "BOM COMMENT", "Status", "Assigned"]
        self.columns_length = len(self.column_names)  # Calculate the columns_length here
        self.create_widgets(self.inner_frame)
        self.window.configure(bg="white")  # set the background color of the window to white
        style = ttk.Style()
        style.configure("TFrame", background="white")  # set the background color of ttk.Frame to white

        self.inner_frame = ttk.Frame(self.canvas, style="TFrame")  # create the inner frame with the updated style

    def create_widgets(self, frame):
        style = ttk.Style()
        style.configure("TButton", font=("Arial", 14, "bold"), width=60, height=2)
        style.map("TButton",
                  foreground=[('active', 'red')],
                  background=[('active', 'blue')])
        style.configure("TButton", background="white")  # Change the button background color to white

        title_label = ttk.Label(frame, text="Welcome Partnership Member!",
                                font=("Arial", 32, "underline"), background="white")
        title_label.pack(pady=10)

        description_label = ttk.Label(frame,
                                      text="This tool allows you to sort your Excel files for our Creation "
                                           "Contact",
                                      font=("Arial", 16, "underline"), background="white")
        description_label.pack(pady=10)

        sort_award_button = ttk.Button(frame, text="Sort Award File", command=self.sort_award_file,
                                       style="TButton")
        sort_award_button.pack(pady=10)

        sort_backlog_button = ttk.Button(frame, text="Sort Backlog File", command=self.sort_backlog_file,
                                         style="TButton")
        sort_backlog_button.pack(pady=10)

        sort_last_ship_date_button = ttk.Button(frame, text="Sort Sales History File",
                                                command=self.sort_by_last_ship_date, style="TButton")
        sort_last_ship_date_button.pack(pady=10)

        sort_ship_and_debit = ttk.Button(frame, text="Sort SND File", command=self.sort_ship_and_debit,
                                         style="TButton")
        sort_ship_and_debit.pack(pady=10)

        sort_vpc = ttk.Button(frame, text="Sort VPC File", command=self.sort_vpc, style="TButton")
        sort_vpc.pack(pady=10)

        add_instructions_for_active_contracts_file = ttk.Label(
            frame,
            text="For the 'Merge Files and Create 'Lost Items' Sheet' button will allow you to merge \n"
                 "your files accordingly now that they are sorted.\n"
                 "Order to Select Files:\n 1. Current Contract\n "
                 "2. Previous Weeks Contract\n 3. Awards File 4. Backlog File\n "
                 "5. Sales History File\n 6. SND File 7. VPC File\n  8. Finally Running File\n"
                 "You will get a success message at the end",
            font=("Arial", 18),
            background="white",
            anchor="center",
            justify="center",
            wraplength=1000
        )
        add_instructions_for_active_contracts_file.pack(pady=2)

        merge_and_create_lost_items_button = ttk.Button(frame, text="Merge Files and Create 'Lost Items' Sheet",
                                                        command=self.merge_files_and_create_lost_items, style="TButton")
        merge_and_create_lost_items_button.pack(pady=10)

        new_instructions = ttk.Label(
            frame,
            text="For the 'Perform VLOOKUP' button \n"
                 "Select the File where you need a Vlookup Completed:\n"
                 "1. Current Active Supplier Contracts File with Data Merged\n",
            font=("Arial", 18),
            background="white",
            anchor="center",
            justify="center",
            wraplength=1000
        )
        new_instructions.pack(pady=10)

        perform_vlookup_button = ttk.Button(frame, text="Perform VLOOKUP",
                                            command=self.perform_vlookup, style="TButton")
        perform_vlookup_button.pack(pady=10)

        logo_label = ttk.Label(frame, background="white")
        logo_label.pack(pady=10)

        logo_image = Image.open('images-videos/Sager-logo.png')
        logo_image = ImageTk.PhotoImage(logo_image)
        logo_label.config(image=logo_image)
        logo_label.image = logo_image

        # Center all the widgets vertically in the frame
        for widget in frame.winfo_children():
            widget.pack_configure(pady=5)

    @staticmethod
    def merge_files_and_create_lost_items():
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

        # Load active and prev contract dataframes
        active_df = pd.read_excel(active_award_file_path, sheet_name=0,
                                  skiprows=1)  # considering headers on 2nd row in 1st file
        prev_contract_file_path = file_paths[1]
        prev_df = pd.read_excel(prev_contract_file_path, sheet_name=1)  # headers on 1st row in 2nd file

        # Always create 'Lost Items' sheet
        lost_items_sheet = active_award_workbook.create_sheet('Lost Items')

        # Add data to 'Lost Items' sheet only if there are lost items
        lost_items_df = prev_df[~prev_df['IPN'].isin(active_df['IPN'])]
        if not lost_items_df.empty:  # check if there are lost items
            for r in dataframe_to_rows(lost_items_df, index=False, header=True):
                lost_items_sheet.append(r)
        else:
            lost_items_sheet.append(list(prev_df.columns))  # append headers only

        # Load and merge data from other files
        for file_path, file_name in zip(file_paths[1:], file_names[1:]):  # Skip active_award_file
            data = pd.read_excel(file_path)

            # Create a new sheet for each file and add data to it
            new_sheet = active_award_workbook.create_sheet(title=file_name)
            for r in dataframe_to_rows(data, index=False, header=True):
                new_sheet.append(r)

        try:
            active_award_workbook.save(active_award_file_path)
            messagebox.showinfo("Success!",
                                "Files merged and 'Lost Items' sheet created successfully with any missing IPN's from last week that are not in the current weeks file.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def browse_files(self):
        self.filename = filedialog.askopenfilename()

    "vlookup not brining in the columns correctly with their matching data to their corresponding IPNS, must fix"

    @staticmethod
    def perform_vlookup():
        try:
            # Ask the user for the contract file paths
            contract_file = filedialog.askopenfilename(title="Select the contract file, where we need a vlookup")

            # Define columns to bring from the reference file
            columns_to_bring = [
                "IPN", "Prev Contract MPN", "Prev Contract Price", "MPN Match", "Price Match MPN",
                "LAST WEEK Contract Change", "Contract Change", "PSoft Part", "count",
                "Corrected PSID Ct", "SUM", "AVG", "DIFF", "PSID All Contract Prices Same?",
                "PS Award Price", "PS Award Exp Date", "PS Awd Cust ID", "Price Match Award",
                "Corp Awd Loaded", "Review Note", "90 DAY PI - NEW PRICE", "PI SENT DATE",
                "DIFF Price Increase", "PI EFF DATE", "12 Month CPN Sales", "GP%", "Cost",
                "Cost Note", "Quote#", "Cost Exp Date", "Cost MOQ", "DIFF LW", "LW Cost",
                "LW Cost Note", "LW Cost Exp Date", "LW Review Note", "Estimated $ Value",
                "Estimated Cost$", "Estimated GP$", "GL-Interconnect Qte - Feb (Y/N)",
                "DS-Battery Qte - Mar (Y/N)", "Part Class", "Sager Stock", "Cost to Use 1",
                "Resale 1", "Price Match", "Sager Min", "Min Match", "New Special Cost",
                "Internal Comments", "New Special Quote#", "SP Exp Date", "Unnamed: 80",
                "Gil Rev Price", "Gil Rev Margin", "Gil Rev MOQ", "Gil Rev SPQ", "Gil Rev Price Match",
                "Price OK", "Min OK", "BOM COMMENT", "Status", "Award Ref #", "PS CPN Award",
                "Product ID", "MFR", "MFR PN", "Award Price", "Release Qty", "Award EAU",
                "Award PO#", "Award Load Date", "Award Price Diff", "Award MOQ Diff", "Award by PSID",
                "Award by PSID Price", "Award by PSID CPN", "Award by PSID MOQ", "MFR.1", "MFR PN.1",
                "Award Price Diff.1", "Award MOQ Diff.1", "MPN Match.1", "RC", "xxx",
                "Current Margin OK for Now?", "PSoft Part.1", "PSID CT", "X", "Quoted Mfg",
                "Quoted Part", "Part Class.1", "**JUNE INT**      Quoted Mfg", "Quoted Part.1",
                "Resale 1.1", "Sager Min.1", "Sager Mult", "Sager NCNR", "Customer Comments",
                "QTE SAME AS CT", "QTE MOQ SAME AS CT", "QTE SPQ SAME AS CT", "PRICE DIFF CURR PS AWARD",
                "AWARD LOADED", "PSoft Part.2", "Last Price Quoted", "Qty Quoted", "Date Quoted",
                "**FEB INT QTE**                                                                   Quoted Mfg",
                "Quoted Part.2", "Qtd Price", "Qtd MOQ", "Qtd SPQ", "NCNR.1", "PSoft Part.3",
                "Feb Qte Price Diff", "Feb Qte MOQ Diff", "Unnamed: 145", "8/14 PS Award CPN",
                "8/14 PS Award PSID", "8/14 PS Award Price", "8/14 PS Award MOQ", "8/14 PS Award Name",
                "Price Diff", "MOQ Diff", "8/14 PS Award End Date"
            ]

            # Load data from 'Prev File' sheet and 'Active Supplier Contracts' sheet
            reference_df = pd.read_excel(contract_file, sheet_name='Prev Contract', header=0)[columns_to_bring]
            contract_df = pd.read_excel(contract_file, sheet_name='Active Supplier Contracts', header=1)

            # # Process the IPN columns
            # reference_df['IPN'] = reference_df['IPN'].astype(str).str.strip().str.upper().str.lstrip('0')
            # contract_df['IPN'] = contract_df['IPN'].astype(str).str.strip().str.upper().str.lstrip('0')

            # Rename the 'Price' column from reference_df
            reference_df = reference_df.rename(columns={'Prev Contract Price': 'Prev_Resale_Price'})

            # Merge on 'IPN'
            final_df = contract_df.merge(reference_df, on='IPN', how='left', suffixes=('', '_y'))

            # Adjust the 'Contract Change' logic
            final_df['Contract Change'] = np.where(final_df['Price'] > final_df['Prev_Resale_Price'], 'Price Increase',
                                                   np.where(final_df['Price'] < final_df['Prev_Resale_Price'],
                                                            'Price Decrease',
                                                            np.where(final_df['Price'] == final_df['Prev_Resale_Price'],
                                                                     'No Change',
                                                                     np.where(pd.isna(final_df['Prev_Resale_Price']),
                                                                              'New Item',
                                                                              'No Change'))))

            # Load all sheets from the contract file
            all_sheets = pd.read_excel(contract_file, sheet_name=None)
            all_sheets['Active Supplier Contracts'] = final_df  # update this sheet with the final_df

            # Ask the user for the output file path
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save the output file as")

            # Write the data to a new Excel file
            if output_file:
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    for sheet_name, df in all_sheets.items():
                        df.to_excel(writer, index=False, sheet_name=sheet_name)

                # Display a success message in a message box
                messagebox.showinfo("Success", "The output file has been saved as: " + output_file)

        except Exception as e:
            messagebox.showerror("Error", str(e))

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
