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
        lost_items_df = pd.read_excel(contract_file, sheet_name='Lost Items')
        awards_df = pd.read_excel(contract_file, sheet_name='Awards')
        snd_df = pd.read_excel(contract_file, sheet_name='SND')
        vpc_df = pd.read_excel(contract_file, sheet_name='VPC')
        backlog_df = pd.read_excel(contract_file, sheet_name='Backlog')
        sales_history_df = pd.read_excel(contract_file, sheet_name='Sales History')

        # Merge on 'IPN' to get the 'PSoft Part' column, these columns being brought in is what we are using for the merge
        active_supplier_df = active_supplier_df.merge(
            prev_contract_df[['IPN', "Price", 'PSoft Part', "Prev Contract MPN", "Prev Contract Price", "MPN Match",
                              "Price Match MPN",
                              "LAST WEEK Contract Change", "Contract Change", "count",
                              "Corrected PSID Ct", "SUM", "AVG", "DIFF", "PSID All Contract Prices Same?",
                              "PS Award Price", "PS Award Exp Date", "PS Awd Cust ID", "Price Match Award",
                              "Corp Awd Loaded", "90 DAY PI - NEW PRICE", "PI SENT DATE",
                              "DIFF Price Increase", "PI EFF DATE", "12 Month CPN Sales", "GP%", "Cost",
                              "Cost Note", "Quote#", "Cost Exp Date", "Cost MOQ", "DIFF LW", "LW Cost",
                              "LW Cost Note", "LW Cost Exp Date", "Review Note"]], on='IPN', how='left')

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
            abs(active_supplier_df['Price_x'] - active_supplier_df['Price_y']) <= tolerance,
            'No Change',
            np.where(active_supplier_df['Price_x'] > active_supplier_df['Price_y'],
                     'Price Increase',
                     np.where(active_supplier_df['Price_x'] < active_supplier_df['Price_y'],
                              'Price Decrease',
                              np.where(pd.isna(active_supplier_df['Price_y']),
                                       'New Item', 'Unknown')))
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

                # This code until the headers_to_color it the formula we incorporated to calculate GP %
                # Its (Resale Price - Cost) / Resale Price
                # Find the columns for 'Price_x', 'Cost', and 'GP%'
                price_x_col = None
                cost_col = None
                gp_col = None

                for col_num, col_cells in enumerate(sheet.columns, start=1):
                    if col_cells[0].value == 'Price_x':
                        price_x_col = col_num
                    elif col_cells[0].value == 'Cost':
                        cost_col = col_num
                    elif col_cells[0].value == 'GP%':
                        gp_col = col_num

                # Check if all the required columns were found
                if price_x_col and cost_col and gp_col:
                    for row in range(2, sheet.max_row + 1):  # Assuming row 1 is the header, so we start from row 2
                        gp_cell = f"{get_column_letter(gp_col)}{row}"
                        price_x_cell = f"{get_column_letter(price_x_col)}{row}"
                        cost_cell = f"{get_column_letter(cost_col)}{row}"
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

# Initial State: The 'Cost' column in the active_supplier_df DataFrame is initially populated
# with whatever values are in the 'Prev Contract' sheet in the 'Cost' column, if there are any.
# This happens as a result of merging active_supplier_df with selected columns from prev_contract_df
# on the 'IPN' column.

# Updating from SND Sheet:

# The code iterates through each row of the active_supplier_df DataFrame.
# For each row, it tries to find a match for the 'PSoft Part' in the snd_df DataFrame using the 'Product ID' column.
# If a match is found and the value from snd_df (specifically the second column, indexed as iloc[0, 1])
# is not NaN (or empty), then the code updates the 'Cost' column of the active_supplier_df with this value.
# If a matching value is found in snd_df, the code then continues to the next row in active_supplier_df
# without checking the vpc_df. This is because the 'SND' sheet has priority; if a 'Cost' value is found there,
# it will be used over any potential match in the 'VPC' sheet.
# Updating from VPC Sheet:
#
# If no match was found in the snd_df or if the matching value was NaN, the code then tries to find a match
# for the 'PSoft Part' in the vpc_df DataFrame using the 'PART ID' column.
# If a match is found and the value from vpc_df (again the second column,
# indexed as iloc[0, 1]) is not NaN (or empty), then the 'Cost' column of the active_supplier_df is
# updated with this value.
# Final State: After iterating through all rows, the 'Cost' column in the active_supplier_df will contain:
#
# The value from the snd_df if a match was found there.
# If no match was found in snd_df or the matched value was NaN, it will contain the
# value from vpc_df if a match was found there.
# If no matches were found in either snd_df or vpc_df (or both had NaN values),
# it will retain the original value from the 'Prev Contract' sheet.
# So, in essence, the 'Cost' column in active_supplier_df is being
# populated with the most recent and relevant data from either snd_df or vpc_df, but
# will retain its original value if no relevant updates are found in those two DataFrames.
