# from tkinter import filedialog, messagebox
# import pandas as pd
# from openpyxl.reader.excel import load_workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
#
#
# def add_running_file_to_workbook(button_to_disable):
#     print("add_running_file_to_workbook called")
#
#     # Ask the user to select the workbook to which the Running File data should be added
#     workbook_file_path = filedialog.askopenfilename(title="Select the file where we need to add the Running File data",
#                                                     filetypes=[("Excel files", "*.xlsx;*.xls")],
#                                                     initialdir="I:\Quotes\Partnership Sales - CM\Creation")
#     if not workbook_file_path:
#         messagebox.showerror("Error", "Workbook selection cancelled.")
#         return
#
#     special_file_name = "Running File - 30 Day Notice Contract Price Increase_Sager - COSTED"
#     running_file_path = filedialog.askopenfilename(title="Select {} file".format(special_file_name),
#                                                    initialdir="I:\Quotes\Partnership Sales - CM\Creation")
#     if not running_file_path:
#         messagebox.showerror("Error", "File selection for Running File cancelled.")
#         return
#
#     # Read the Running File data into a DataFrame
#     running_data = pd.read_excel(running_file_path)
#
#     # Load the selected workbook
#     selected_workbook = load_workbook(workbook_file_path)
#     running_sheet = selected_workbook.create_sheet(title=special_file_name)
#
#     # Transfer data from DataFrame to the selected workbook
#     for r in dataframe_to_rows(running_data, index=False, header=True):
#         running_sheet.append(r)
#
#     try:
#         selected_workbook.save(workbook_file_path)
#         messagebox.showinfo("Success!", "Running File data added successfully.")
#         button_to_disable.config(state="disabled")
#     except Exception as e:
#         messagebox.showerror("Error", "Failed to add Running File. Error: " + str(e))
