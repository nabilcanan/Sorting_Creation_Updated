import tkinter as tk
import tkinter.ttk as ttk
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
import numpy as np
import pandas as pd
from PIL import ImageTk, Image
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from tkinter import filedialog, messagebox
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


class ExcelSorter:
    def __init__(self):
        self.run_queries_class = None
        self.filename = None
        self.window = tk.Tk()
        self.window.title("Sorting Creation Files For Contract")
        self.window.configure(bg="white")
        self.window.geometry("880x600")  # Usually 600 for normal wundow size

        # Create a canvas and a vertical scrollbar
        self.canvas = tk.Canvas(self.window)
        self.scrollbar = ttk.Scrollbar(self.window, orient="vertical", command=self.canvas.yview)

        # Configure the canvas to respond to the scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a frame to hold your widgets, and add it to the canvas
        self.inner_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((self.window.winfo_width() / 2, 0), window=self.inner_frame, anchor="n")

        # Configure the canvaZs's scroll-region to encompass the frame
        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Pack the scrollbar, making sure it sticks to the right side
        self.scrollbar.pack(side="right", fill="y")

        # Configure the canvas to expand and fill the window
        self.canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)

        # Canvas - Scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.canvas.yview)

        self.file_paths = []
        self.create_widgets(self.inner_frame)
        self.window.configure(bg="white")  # set the background color of the window to white
        style = ttk.Style()
        style.configure("TFrame", background="white")  # set the background color of ttk.Frame to white

        self.inner_frame = ttk.Frame(self.canvas, style="TFrame")  # create the inner frame with the updated style

    def create_widgets(self, frame):
        style = ttk.Style()
        style.configure("TButton", font=("Times New Roman", 16, "bold"), width=60, height=2)
        style.map("TButton",
                  foreground=[('active', 'red')],
                  background=[('active', 'blue')])
        style.configure("TButton", background="white")  # Change the button background color to white

        title_label = ttk.Label(frame, text="Welcome Partnership Member!",
                                font=("Times New Roman", 32, "underline"), background="white", foreground="#103d81")
        title_label.pack(pady=10)

        description_label = ttk.Label(frame,
                                      text="This tool allows you to sort your Excel files for our Creation "
                                           "Contact",
                                      font=("Times New Roman", 16, "underline"), background="white")
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
            font=("Times New Roman", 19),
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
                 "1. Select the file where you now need your vlookup completed.\n"
                 "(This is the same file where all your files are now merged.) ",
            font=("Times New Roman", 19),
            background="white",
            anchor="center",
            justify="center",
            wraplength=1000
        )
        new_instructions.pack(pady=10)

        perform_vlookup_button = ttk.Button(frame, text="Perform VLOOKUP to new file",
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
        active_df = pd.read_excel(active_award_file_path, sheet_name=0, skiprows=1)  # considering headers on 2nd row

        prev_contract_file_path = file_paths[1]
        prev_df = pd.read_excel(prev_contract_file_path, sheet_name=0)  # headers on 1st row

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
                                "Files merged and 'Lost Items' sheet created "
                                "successfully with any missing IPN's from last week that "
                                "are not in the current weeks file.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    "vlookup not brining in the columns correctly with their matching data to their corresponding IPNS, must fix"

    @staticmethod
    def perform_vlookup():
        try:
            # Ask the user for the contract file paths
            contract_file = filedialog.askopenfilename(title="Select the contract file, where we need a vlookup")

            # Define columns to bring from the reference file
            columns_to_bring = [
                "IPN", "CM", "Item", "Price", "Prev Contract MPN", "Prev Contract Price", "MPN Match",
                "Price Match MPN",
                "LAST WEEK Contract Change", "Contract Change", "PSoft Part", "count",
                "Corrected PSID Ct", "SUM", "AVG", "DIFF", "PSID All Contract Prices Same?",
                "PS Award Price", "PS Award Exp Date", "PS Awd Cust ID", "Price Match Award",
                "Corp Awd Loaded", "90 DAY PI - NEW PRICE", "PI SENT DATE",
                "DIFF Price Increase", "PI EFF DATE", "12 Month CPN Sales", "GP%", "Cost",
                "Cost Note", "Quote#", "Cost Exp Date", "Cost MOQ", "DIFF LW", "LW Cost",
                "LW Cost Note", "LW Cost Exp Date", "Review Note"]

            # Load data from 'Prev File' sheet and 'Active Supplier Contracts' sheet
            reference_df = pd.read_excel(contract_file, sheet_name='Prev Contract', header=0)[columns_to_bring]
            print("Headers in reference_df:", reference_df.columns.tolist())  # Print headers of reference_df

            contract_df = pd.read_excel(contract_file, sheet_name='Active Supplier Contracts', header=1)
            print("Headers in contract_df:", contract_df.columns.tolist())  # Print headers of contract_df

            # Rename the 'Price' column from reference_df
            reference_df = reference_df.rename(columns={'Prev Contract Price': 'Prev_Resale_Price'})

            # Merge on 'IPN'
            final_df = contract_df.merge(reference_df, on='IPN', how='left', suffixes=('', '_y'))
            print("Headers in final_df:", final_df.columns.tolist())  # Print headers of final_df

            # This was for getting the GP column in as a percent, did not convert the way
            # I needed to, so leaving it out for now
            # final_df['GP%'] = (final_df['GP%'] * 100).astype(str) + '%'

            tolerance = 0.0001

            final_df['Contract Change'] = np.where(abs(final_df['Price'] - final_df['Price_y']) <= tolerance,
                                                   'No Change',
                                                   np.where(final_df['Price'] > final_df['Price_y'],
                                                            'Price Increase',
                                                            np.where(final_df['Price'] < final_df['Price_y'],
                                                                     'Price Decrease',
                                                                     np.where(pd.isna(final_df['Price_y']),
                                                                              'New Item', 'Unknown'))))

            # Load all sheets from the contract file
            all_sheets = pd.read_excel(contract_file, sheet_name=None)
            all_sheets['Active Supplier Contracts'] = final_df  # update this sheet with the final_df

            # Remove the unwanted sheet
            if '_SettingsCurrency' in all_sheets:
                del all_sheets['_SettingsCurrency']

            # Ask the user for the output file path
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save the output file as")

            # Write the data to a new Excel file
            if output_file:
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    for sheet_name, df in all_sheets.items():
                        df.to_excel(writer, index=False, sheet_name=sheet_name)

                    ws = writer.sheets['Active Supplier Contracts']

                    # Make headers wraptext
                    for cell in ws["1:1"]:  # This specifies the first row, which are the headers
                        cell.alignment = Alignment(wrap_text=True)

                    # Map headers to their respective colors, these will all be diff colors
                    headers_to_color = {
                        'GP%': "0000FFFF",
                        'Cost': "0000FFFF",
                        'Cost Note': "0000FFFF",
                        'Quote#': "0000FFFF",
                        'Cost Exp Date': "0000FFFF",
                        'Cost MOQ': "0000FFFF",
                        'IPN': "00FFFF00",
                        'MPN': "00FFFF00",
                        'MFG': "00FFFF00",
                        'Customer Name': "00FFFF00",
                    }

                    for row in ws.iter_rows(min_row=1, max_row=1):
                        for cell in row:
                            if cell.value in headers_to_color:
                                fill_color = headers_to_color[cell.value]
                                fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                                cell.fill = fill

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
