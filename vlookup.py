from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.styles import PatternFill
import numpy as np
from openpyxl.utils import get_column_letter


def perform_vlookup():
    try:
        # Ask the user for the contract file paths
        contract_file = filedialog.askopenfilename(title="Select the contract file, where we need a vlookup",
                                                   initialdir="I:\Quotes\Partnership Sales - CM\Creation")

        # Load all the sheets, these are being loaded for the vlookup function and updating our SND and VPC Cost's
        active_supplier_df = pd.read_excel(contract_file, sheet_name='Active Supplier Contracts', header=1)
        prev_contract_df = pd.read_excel(contract_file, sheet_name='Prev Contract', header=0)

        # Check if 'Price' exists in prev_contract_df
        if 'Price' not in prev_contract_df.columns:
            raise ValueError("'Price' column not found in the prev_contract_df DataFrame.")

        # Merge using 'Price' from prev_contract_df and rename it to 'Price_x' in the resulting dataframe
        active_supplier_df = active_supplier_df.merge(prev_contract_df[['IPN', 'Price']], on='IPN', how='left').rename(
            columns={'Price': 'Price_x'})

        # After this merge, active_supplier_df should have 'Price_x'. Validate this:
        if 'Price_x' not in active_supplier_df.columns:
            raise ValueError("'Price_x' column was not successfully merged into active_supplier_df.")

        lost_items_df = pd.read_excel(contract_file, sheet_name='Lost Items')
        awards_df = pd.read_excel(contract_file, sheet_name='Awards')
        snd_df = pd.read_excel(contract_file, sheet_name='SND')
        vpc_df = pd.read_excel(contract_file, sheet_name='VPC')
        backlog_df = pd.read_excel(contract_file, sheet_name='Backlog')
        sales_history_df = pd.read_excel(contract_file, sheet_name='Sales History')

        # Merge on 'IPN' to get the 'PSoft Part' column, these columns being brought in is what we are using for the

        # merge
        active_supplier_df = active_supplier_df.merge(
            prev_contract_df[
                ['IPN', "Price", 'PSoft Part', "Prev Contract MPN", "Prev Contract Price", "MPN Match",
                 "Price Match MPN",
                 "LAST WEEK Contract Change", "Contract Change", "count",
                 "Corrected PSID Ct", "SUM", "AVG", "DIFF", "PSID All Contract Prices Same?",
                 "PS Award Price", "PS Award Exp Date", "PS Awd Cust ID", "Price Match Award",
                 "Corp Awd Loaded", "90 DAY PI - NEW PRICE", "PI SENT DATE",
                 "DIFF Price Increase", "PI EFF DATE", "12 Month CPN Sales", "GP%", "Cost",
                 "Cost Note", "Quote#", "Cost Exp Date", "Cost MOQ",
                 "Review Note"]], on='IPN', how='left')

        # Iterate through each row in the active_supplier_df to look for a match in SND and VPC
        for idx, row in active_supplier_df.iterrows():
            psoft_part = row['PSoft Part']

            # Check SND using the 'Product ID' column
            matching_snd = snd_df[snd_df['Product ID'] == psoft_part]
            if not matching_snd.empty and not pd.isna(matching_snd.iloc[0, 1]):
                active_supplier_df.at[idx, 'Cost'] = matching_snd.iloc[0, 1]
                continue  # If found in SND, skip checking VPC for the same PSoft Part

            # Check VPC using the 'PART ID' column
            matching_vpc = vpc_df[vpc_df['PART ID'] == psoft_part]
            if not matching_vpc.empty and not pd.isna(matching_vpc.iloc[0, 1]):
                active_supplier_df.at[idx, 'Cost'] = matching_vpc.iloc[0, 1]

        print(active_supplier_df.columns)

        tolerance = 0.01  # you can set it to any value you deem fit

        active_supplier_df['Contract Change'] = np.where(
            abs(active_supplier_df['Price'] - active_supplier_df['Price_x']) <= tolerance,
            'No Change',
            np.where(active_supplier_df['Price'] > active_supplier_df['Price_x'],
                     'Price Increase',
                     np.where(active_supplier_df['Price'] < active_supplier_df['Price_x'],
                              'Price Decrease',
                              'New Item'))
        )

        # Ask the user for the output file path
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save the output file as",
                                                   initialdir="I:\Quotes\Partnership Sales - CM\Creation")

        # Write all the DataFrames to the new Excel file in the specified order
        if output_file:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                active_supplier_df.to_excel(writer, index=False, sheet_name='Active Supplier Contracts')
                prev_contract_df.to_excel(writer, index=False, sheet_name='Prev Contract')
                lost_items_df.to_excel(writer, index=False, sheet_name='Lost Items')
                awards_df.to_excel(writer, index=False, sheet_name='Awards')
                snd_df.to_excel(writer, index=False, sheet_name='SND')
                vpc_df.to_excel(writer, index=False, sheet_name='VPC')
                backlog_df.to_excel(writer, index=False, sheet_name='Backlog')
                sales_history_df.to_excel(writer, index=False, sheet_name='Sales History')

                # Grabbing the workbook and the desired sheet
                workbook = writer.book
                sheet = workbook['Active Supplier Contracts']

                # Define the columns for 'Price_x', 'Cost', 'GP%', 'Cost Exp Date', 'Award Date', and 'Last Update Date'
                price_x_col, cost_col, gp_col, date_col, award_date_col, last_update_date_col = None, None, None, None, None, None

                for col_num, col_cells in enumerate(sheet.columns, start=1):
                    col_val = col_cells[0].value  # header value in current column
                    if col_val == 'Price_x':
                        price_x_col = col_num
                    elif col_val == 'Cost':
                        cost_col = col_num
                    elif col_val == 'GP%':
                        gp_col = col_num
                    elif col_val == 'Cost Exp Date':
                        date_col = col_num
                    elif col_val == 'Award Date':
                        award_date_col = col_num
                    elif col_val == 'Last Update Date':
                        last_update_date_col = col_num

                # Check if all the required columns were found and apply formatting
                if all([price_x_col, cost_col, gp_col, date_col, award_date_col, last_update_date_col]):
                    for row in range(2, sheet.max_row + 1):  # Assuming row 1 is the header, so we start from row 2
                        gp_cell = f"{get_column_letter(gp_col)}{row}"
                        price_x_cell = f"{get_column_letter(price_x_col)}{row}"
                        cost_cell = f"{get_column_letter(cost_col)}{row}"
                        date_cell = f"{get_column_letter(date_col)}{row}"
                        award_date_cell = f"{get_column_letter(award_date_col)}{row}"
                        last_update_date_cell = f"{get_column_letter(last_update_date_col)}{row}"

                        # Format the cells
                        sheet[gp_cell].number_format = '0.00%'  # GP% as percent
                        sheet[cost_cell].number_format = '$0.0000'  # Cost as dollar with four decimal places
                        sheet[date_cell].number_format = 'MM/DD/YYYY'  # Date as MM/DD/YYYY
                        sheet[award_date_cell].number_format = 'MM/DD/YYYY'  # Award Date as MM/DD/YYYY
                        sheet[last_update_date_cell].number_format = 'MM/DD/YYYY'  # Last Update Date as MM/DD/YYYY

                        # Apply formula to GP%
                        formula = f"=IF({price_x_cell}=0,0,({price_x_cell} - {cost_cell}) / {price_x_cell})"
                        sheet[gp_cell] = formula

                headers_to_color = {
                    'Price': "0000FFFF",
                    'GP%': "0000FFFF",
                    'Cost': "0000FFFF",
                    'Cost Note': "0000FFFF",
                    'Quote#': "0000FFFF",
                    'Cost Exp Date': "0000FFFF",
                    'Cost MOQ': "0000FFFF",
                    'PSoft Part': "00FFFF00",
                    'MPN': "0000FFFF",
                    'MFG': "0000FFFF",
                    'EAU': "0000FFFF",
                    'MOQ': "0000FFFF",
                    'MPQ': "0000FFFF",
                    'NCNR': "0000FFFF"
                }

                for col_num, col_cells in enumerate(sheet.columns, start=1):
                    if col_cells[0].value in headers_to_color:
                        col_cells[0].fill = PatternFill(start_color=headers_to_color[col_cells[0].value],
                                                        end_color=headers_to_color[col_cells[0].value],
                                                        fill_type="solid")

            # Display a success message in a message box
            messagebox.showinfo("Success", "The output file has been saved as: " + output_file)

    except Exception as e:
        messagebox.showerror("Error Process was Cancelled", str(e))
