from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def merge_files_and_create_lost_items():
    print("merge_files_and_create_lost_items called")
    file_names = ["Active Contract File", "Prev Contract", "Awards", "Backlog", "Sales History", "SND", "VPC"]
    file_paths = []

    for file_name in file_names:
        file_path = filedialog.askopenfilename(title="Select {} file".format(file_name),
                                               initialdir="I:\Quotes\Partnership Sales - CM\Creation")
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


def add_running_file_to_workbook():
    print("add_running_file_to_workbook called")

    # Ask the user to select the workbook to which the Running File data should be added
    workbook_file_path = filedialog.askopenfilename(title="Select the workbook to add the Running File data",
                                                    filetypes=[("Excel files", "*.xlsx;*.xls")],
                                                    initialdir="I:\Quotes\Partnership Sales - CM\Creation")
    if not workbook_file_path:
        messagebox.showerror("Error", "Workbook selection cancelled.")
        return

    special_file_name = "Running File - 30 Day Notice Contract Price Increase_Sager - COSTED"
    running_file_path = filedialog.askopenfilename(title="Select {} file".format(special_file_name),
                                                   initialdir="I:\Quotes\Partnership Sales - CM\Creation")
    if not running_file_path:
        messagebox.showerror("Error", "File selection for Running File cancelled.")
        return

    # Load the selected workbook and add the running file data
    selected_workbook = load_workbook(workbook_file_path)
    running_data = pd.read_excel(running_file_path)
    running_sheet = selected_workbook.create_sheet(title=special_file_name)
    for r in dataframe_to_rows(running_data, index=False, header=True):
        running_sheet.append(r)

    try:
        selected_workbook.save(workbook_file_path)
        messagebox.showinfo("Success!", "Running File data added successfully.")
    except Exception as e:
        messagebox.showerror("Error", "Failed to add Running File. Error: " + str(e))
