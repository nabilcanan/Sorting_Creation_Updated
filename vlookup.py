from datetime import datetime
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.styles import PatternFill, Alignment
import numpy as np
from openpyxl.utils import get_column_letter
from colored_headers import headers_to_color


def perform_vlookup(button_to_disable):
    try:
        # Ask the user for the contract file
        contract_file = filedialog.askopenfilename(title="Select the contract file, where we need a vlookup",
                                                   initialdir="P:\Partnership_Python_Projects\Creation\test_001")

        # Load 'Active Supplier Contracts' and 'Prev Contract' sheets
        active_supplier_df = pd.read_excel(contract_file, sheet_name='Active Supplier Contracts', header=1)
        prev_contract_df = pd.read_excel(contract_file, sheet_name='Prev Contract', header=0)
        lost_items_df = pd.read_excel(contract_file, sheet_name='Lost Items')
        awards_df = pd.read_excel(contract_file, sheet_name='Awards')
        snd_df = pd.read_excel(contract_file, sheet_name='SND')
        vpc_df = pd.read_excel(contract_file, sheet_name='VPC')
        backlog_df = pd.read_excel(contract_file, sheet_name='Backlog')
        sales_history_df = pd.read_excel(contract_file, sheet_name='Sales History')

        print("Loaded 'Active Supplier Contracts' sheet with shape:", active_supplier_df.shape)
        print("Loaded 'Prev Contract' sheet with shape:", prev_contract_df.shape)

        # Drop the 'LW PRICE' column if it exists in the 'Prev Contract' dataframe
        if 'LW PRICE' in prev_contract_df.columns:
            prev_contract_df.drop('LW PRICE', axis=1, inplace=True)

        # Rename the 'Price' column from 'Prev Contract' dataframe to 'LW PRICE'
        prev_contract_df.rename(columns={'Price': 'LW PRICE'}, inplace=True)

        # merge
        active_supplier_df = active_supplier_df.merge(
            prev_contract_df[
                ['IPN', 'LW PRICE', 'PSoft Part', "Prev Contract MPN", "MPN Match",
                 "Price Match MPN",
                 "Contract Change", "count",
                 "Corrected PSID Ct", "SUM", "AVG", "DIFF", "PSID All Contract Prices Same?",
                 "PS Award Price", "PS Award Exp Date", "PS Awd Cust ID", "Price Match Award",
                 "Corp Awd Loaded", "90 DAY PI - NEW PRICE", "PI SENT DATE",
                 "DIFF Price Increase", "PI EFF DATE", "12 Month CPN Sales", "GP%", "Cost",
                 "Cost Note", "Quote#", "Cost Exp Date", "Cost MOQ",
                 "Review Note", "LW Cost", "LW Quote#", "LW Cost Exp Date", "LW Review Note"]], on='IPN', how='left')

        # Calculate the counts for each 'PSoft Part'
        psoft_part_counts = active_supplier_df['PSoft Part'].value_counts()

        # Update the 'count' column in active_supplier_df with the new counts
        active_supplier_df['count'] = active_supplier_df['PSoft Part'].map(psoft_part_counts)

        # Iterate through each row in the active_supplier_df to look for a match in SND and VPC
        for idx, row in active_supplier_df.iterrows():
            psoft_part = row['PSoft Part']
            ipn = row['IPN']

            # Check SND using the 'Product ID' column
            matching_snd = snd_df[snd_df['Product ID'] == psoft_part]
            if not matching_snd.empty and not pd.isna(matching_snd.iloc[0, 1]):
                active_supplier_df.at[idx, 'Cost'] = matching_snd.iloc[0, 1]
                continue  # If found in SND, skip checking VPC for the same PSoft Part

            # Check VPC using the 'PART ID' column
            matching_vpc = vpc_df[vpc_df['PART ID'] == psoft_part]
            if not matching_vpc.empty and not pd.isna(matching_vpc.iloc[0, 1]):
                active_supplier_df.at[idx, 'Cost'] = matching_vpc.iloc[0, 1]

            # Get the column index for 'End Date'
            end_date_col_index = awards_df.columns.get_loc('End Date')

            # Check Awards ex p date using the 'Award CPN' column from the awards_df we loaded then perform the action
            matching_awards = awards_df[awards_df['Award CPN'] == ipn]
            if not matching_awards.empty and not pd.isna(matching_awards.iloc[0, end_date_col_index]):
                active_supplier_df.at[idx, 'PS Award Exp Date'] = matching_awards.iloc[0, end_date_col_index]

                # Check Awards for 'Award Price' and 'Award Cust ID' using the 'Award CPN' column from awards_df
                matching_awards = awards_df[awards_df['Award CPN'] == ipn]
                if not matching_awards.empty:
                    # Update 'PS Award Price'
                    if pd.notna(matching_awards['Award Price'].iloc[0]):
                        active_supplier_df.at[idx, 'PS Award Price'] = matching_awards['Award Price'].iloc[0]

                    # Update 'PS AWD CUST ID' from awards_df
                    if pd.notna(matching_awards['Award Cust ID'].iloc[0]):
                        active_supplier_df.at[idx, 'PS Awd Cust ID'] = matching_awards['Award Cust ID'].iloc[0]

        print(active_supplier_df.columns)

        # The Contract Change comparison is done between 'Price' and 'LW PRICE'.
        # tolerance = 0.01  # you can set it to any value you deem fit

        print("Shape of active_supplier_df['Price']:", active_supplier_df['Price'].shape)
        print("Shape of active_supplier_df['LW PRICE']:", active_supplier_df['LW PRICE'].shape)

        active_supplier_df['Contract Change'] = np.where(
            active_supplier_df['LW PRICE'].isna(),  # Check if 'LW PRICE' is NaN or null
            'New Item',
            np.where(
                active_supplier_df['Price'] == active_supplier_df['LW PRICE'],
                'No Change',
                np.where(
                    active_supplier_df['Price'] > active_supplier_df['LW PRICE'],
                    'Price Increase',
                    'Price Decrease'  # Since we've covered all other scenarios, this can be the else condition.
                )
            )
        )

        # Ask the user for the output file path
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save the output file as",
                                                   initialdir="P:\Partnership_Python_Projects\Creation\test_001")

        # Write all the DataFrames to the new Excel file in the specified order
        if output_file:
            with (pd.ExcelWriter(output_file, engine='openpyxl') as writer):
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

                # Freeze the top row
                sheet.freeze_panes = 'A2'

                # Freeze the column H2
                sheet.freeze_panes = "H2"

                # Turn on filters for the top row only
                sheet.auto_filter.ref = sheet.dimensions

                # Wrap text for the first row, makes it look neater
                for cell in sheet["1:1"]:
                    cell.alignment = Alignment(wrap_text=True)

                # Define the columns for 'Price_x', 'Cost', 'GP%', 'Cost Exp Date', 'Award Date', and 'Last Update Date'
                price_x_col, cost_col, gp_col, date_col, award_date_col, last_update_date_col, pi_sent_date_col = None, None, None, None, None, None, None

                for col_num, col_cells in enumerate(sheet.columns, start=1):
                    col_val = col_cells[0].value  # header value in current column
                    if col_val == 'Price':
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
                    elif col_val == 'PI SENT DATE':
                        pi_sent_date_col = col_num

                    # Formatting for the "PI SENT DATE" column
                    if pi_sent_date_col:
                        for row in range(2, sheet.max_row + 1):
                            pi_sent_date_cell = f"{get_column_letter(pi_sent_date_col)}{row}"

                            cell_value = sheet[pi_sent_date_cell].value

                            # If cell value is '-2' or is an out-of-range date
                            if cell_value == "-2" or str(cell_value) == "####" or cell_value is None:
                                sheet[pi_sent_date_cell].value = ""
                            else:
                                try:
                                    # Check if the value can be interpreted as a date
                                    parsed_date = datetime.strptime(str(cell_value), "%Y-%m-%d %H:%M:%S")
                                    if parsed_date.year < 1900 or parsed_date.year > 9999:  # Excel's date limits
                                        sheet[pi_sent_date_cell].value = ""
                                except ValueError:
                                    sheet[pi_sent_date_cell].value = ""

                    # Now apply the date format
                    if pi_sent_date_col:
                        for row in range(2, sheet.max_row + 1):
                            pi_sent_date_cell = f"{get_column_letter(pi_sent_date_col)}{row}"
                            sheet[pi_sent_date_cell].number_format = 'MM/DD/YYYY'  # Format as MM/DD/YYYY

                # Check if all the required columns were found and apply formatting
                if all([price_x_col, cost_col, gp_col, date_col, award_date_col, last_update_date_col,
                        pi_sent_date_col]):
                    for row in range(2, sheet.max_row + 1):  # Assuming row 1 is the header, so we start from row 2
                        gp_cell = f"{get_column_letter(gp_col)}{row}"
                        price_x_cell = f"{get_column_letter(price_x_col)}{row}"
                        cost_cell = f"{get_column_letter(cost_col)}{row}"
                        date_cell = f"{get_column_letter(date_col)}{row}"
                        award_date_cell = f"{get_column_letter(award_date_col)}{row}"
                        last_update_date_cell = f"{get_column_letter(last_update_date_col)}{row}"
                        pi_sent_date_cell = f"{get_column_letter(pi_sent_date_col)}{row}"

                        # Format the cells
                        sheet[price_x_cell].number_format = '$0.0000'  # Formats the Price cells accordingly
                        sheet[gp_cell].number_format = '0.00%'  # GP% as percent
                        sheet[cost_cell].number_format = '$0.0000'  # Cost as dollar with four decimal places
                        sheet[date_cell].number_format = 'MM/DD/YYYY'  # Date as MM/DD/YYYY
                        sheet[award_date_cell].number_format = 'MM/DD/YYYY'  # Award Date as MM/DD/YYYY
                        sheet[last_update_date_cell].number_format = 'MM/DD/YYYY'  # Last Update Date as MM/DD/YYYY
                        sheet[pi_sent_date_cell].number_format = 'MM/DD/YYYY'  # Format as MM/DD/YYYY

                        # Apply formula to GP%
                        formula = f"=IF({price_x_cell}=0,0,({price_x_cell} - {cost_cell}) / {price_x_cell})"
                        sheet[gp_cell] = formula

                for col_num, col_cells in enumerate(sheet.columns, start=1):
                    if col_cells[0].value in headers_to_color:
                        col_cells[0].fill = PatternFill(start_color=headers_to_color[col_cells[0].value],
                                                        end_color=headers_to_color[col_cells[0].value],
                                                        fill_type="solid")

            # Display a success message in a message box
            messagebox.showinfo("Success", "The output file has been saved as: " + output_file)
            button_to_disable.config(state="disabled")

    except Exception as e:
        messagebox.showerror("Error Process was Cancelled", str(e))
