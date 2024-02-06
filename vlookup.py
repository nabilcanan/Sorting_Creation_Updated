from datetime import datetime, timedelta
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.styles import PatternFill, Alignment
import numpy as np
from openpyxl.utils import get_column_letter
from colored_headers import headers_to_color
import psutil


def perform_vlookup(button_to_disable):
    try:
        # ------------------ SELECT CONTRACT FILE -------------------------------
        # Ask the user for the contract file
        contract_file = filedialog.askopenfilename(title="Select the contract file, where we need a vlookup",
                                                   initialdir="P:\Partnership_Python_Projects\Creation\test_001")

        # ------------------ LOAD DATAFRAMES ------------------------------------
        # Load 'Active Supplier Contracts' and 'Prev Contract' sheets
        active_supplier_df = pd.read_excel(contract_file, sheet_name='Active Supplier Contracts', header=1)
        prev_contract_df = pd.read_excel(contract_file, sheet_name='Prev Contract', header=0, dtype={'PSoft Part': str})
        lost_items_df = pd.read_excel(contract_file, sheet_name='Lost Items')
        awards_df = pd.read_excel(contract_file, sheet_name='Awards')
        snd_df = pd.read_excel(contract_file, sheet_name='SND')
        vpc_df = pd.read_excel(contract_file, sheet_name='VPC')
        backlog_df = pd.read_excel(contract_file, sheet_name='Backlog')
        sales_history_df = pd.read_excel(contract_file, sheet_name='Sales History')

        print("Loaded 'Active Supplier Contracts' sheet with shape:", active_supplier_df.shape)
        print("Loaded 'Prev Contract' sheet with shape:", prev_contract_df.shape)
        # ---------------- End of Loading Sheets ---------------------------------

        # ------------------ PRELIMINARY DATA PREPARATION ----------------------------
        # Drop the 'LW PRICE' column if it exists in the 'Prev Contract' dataframe
        if 'LW PRICE' in prev_contract_df.columns:
            prev_contract_df.drop('LW PRICE', axis=1, inplace=True)

        # Rename the 'Price' column from 'Prev Contract' dataframe to 'LW PRICE'
        prev_contract_df.rename(columns={'Price': 'LW PRICE'}, inplace=True)

        # We will add this in after we discuss in our meeting, basically we are going into the prev contract
        # and bringing in the Cost column and brining it into the LW Cost column the yellow one at the end of
        # the active dataframe
        # # ------------------ MERGE 'LW COST' FROM PREV CONTRACT TO ACTIVE SUPPLIER DF ------------------
        # if 'Cost' in prev_contract_df.columns:
        #     prev_contract_df.rename(columns={'Cost': 'LW Cost'}, inplace=True)
        #     lw_cost_mapping = prev_contract_df.set_index('IPN')['LW Cost'].to_dict()
        #     active_supplier_df['LW Cost'] = active_supplier_df['IPN'].map(lw_cost_mapping)
        # # ------------------ Enb of MERGE 'LW COST' FROM PREV CONTRACT TO ACTIVE SUPPLIER DF ------------

        # ---------------- Adding the merge algorithm to bring in columns from prev contract dataframe --------------
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
        # ---------------- End of Adding the merge algorithm to bring in columns from prev contract dataframe -------

        #  ------Calculate the counts for each 'PSoft Part' and display them with the value number we have ----------
        psoft_part_counts = active_supplier_df['PSoft Part'].value_counts()
        active_supplier_df['count'] = active_supplier_df['PSoft Part'].map(psoft_part_counts)

        active_supplier_df['IPN'] = active_supplier_df['IPN'].astype(str).str.strip()
        active_supplier_df['PSoft Part'] = active_supplier_df['PSoft Part'].astype(str).str.strip()
        awards_df.columns = awards_df.columns.str.strip()
        #  ------ End of Calculate the counts for each 'PSoft Part' and display them with the value number we have --

        # ------- This for loop is used for the next coming if statements and adjusting from other df ---------------
        for idx, row in active_supplier_df.iterrows():
            ipn = row['IPN']  # we use the IPN value to go back and forth between the awards dataframe
            psoft_part = row['PSoft Part']

            # ------------------ BACKLOG VALUE MAPPING TO LOST ITEMS ------------------
            # Create a mapping from 'Backlog CPN' to 'Backlog Value' in backlog_df
            backlog_value_mapping = backlog_df.set_index('Backlog CPN')['Backlog Value'].to_dict()
            # Apply the mapping to 'IPN' column in lost_items_df to get 'Backlog Value'
            lost_items_df['Backlog Value'] = lost_items_df['IPN'].map(backlog_value_mapping)
            # Format the 'Backlog Value' column as currency with 2 decimal places
            lost_items_df['Backlog Value'] = lost_items_df['Backlog Value'].apply(
                lambda x: "" if pd.isnull(x) or x == 0 else "${:,.2f}".format(x))
            # ---------------------------------------------------------------------------

            # ------------------- 12 Month CPN Sales Column Logic ----------------------
            # Example column name: 'YourDateColumnName'
            date_column = 'Last Ship Date'  # Replace with your actual date column name
            # Convert the date column to datetime format
            sales_history_df[date_column] = pd.to_datetime(sales_history_df[date_column], errors='coerce')
            # Calculate the date for 12 months ago
            one_year_ago = datetime.now() - timedelta(days=365)
            # Filter sales_history_df to include only the last 12 months
            sales_history_df_filtered = sales_history_df[sales_history_df[date_column] >= one_year_ago]
            # Continue with your grouping and summing logic as before
            sales_history_grouped = sales_history_df_filtered.groupby('Last Ship CPN')['Net'].sum().reset_index()
            # Map these summed 'NET' values to '12 Month CPN Sales' in active_supplier_df
            # Convert the grouped data into a dictionary for easy mapping
            sales_net_mapping = sales_history_grouped.set_index('Last Ship CPN')['Net'].to_dict()
            # Map the summed 'NET' values to 'active_supplier_df' based on 'IPN'
            active_supplier_df['12 Month CPN Sales'] = active_supplier_df['IPN'].map(sales_net_mapping)
            # -----------------------------------------------------------------------------------

            # ------------ SND Cost updates Logic -----------------------------------------
            # Check SND using the 'Product ID' column
            matching_snd = snd_df[snd_df['Product ID'] == psoft_part]
            if not matching_snd.empty and not pd.isna(matching_snd.iloc[0, 1]):
                active_supplier_df.at[idx, 'Cost'] = matching_snd.iloc[0, 1]
                continue  # If found in SND, skip checking VPC for the same PSoft Part
            # -----------------------------------------------------------------------------

            # ------------ VPC Cost updates Logic -----------------------------------------
            # Check VPC using the 'PART ID' column
            matching_vpc = vpc_df[vpc_df['PART ID'] == psoft_part]
            if not matching_vpc.empty and not pd.isna(matching_vpc.iloc[0, 1]):
                active_supplier_df.at[idx, 'Cost'] = matching_vpc.iloc[0, 1]
            # -----------------------------------------------------------------------------

            # ------------------ UPDATE 'CORP AWARD LOADED' STATUS ------------------
            # Normalize 'Award CPN' values from awards_df for a more accurate lookup (trim spaces, ensure consistent case)
            normalized_award_cpn_set = set(awards_df['Award CPN'].str.strip().str.lower())

            # Update 'Corp Award Loaded' based on the presence of a normalized 'IPN' in the normalized award CPN set
            # Use a different variable name inside the lambda to avoid shadowing
            active_supplier_df['Corp Awd Loaded'] = active_supplier_df['IPN'].str.strip().str.lower().apply(
                lambda x: 'Y' if x in normalized_award_cpn_set else 'N'
            )
            # ------------------ End of Corp Award loaded status --------------------

            # ------------------ PRICE MATCH CHECK BETWEEN ACTIVE SUPPLIER AND AWARDS DATAFRAME ------------------
            # Convert prices to a consistent type (e.g., float) and round to a certain decimal precision if needed
            active_supplier_df['PS Award Price'] = active_supplier_df['PS Award Price'].apply(pd.to_numeric,
                                                                                              errors='coerce')
            awards_df['Award Price'] = awards_df['Award Price'].apply(pd.to_numeric, errors='coerce')

            # Create a dictionary for 'Award CPN' and their 'Award Price' from awards_df
            award_price_mapping = awards_df.set_index('Award CPN')['Award Price'].to_dict()

            # Perform the comparison and update 'Price Match Award'
            active_supplier_df['Price Match Award'] = active_supplier_df.apply(
                lambda x: 'Y' if np.isclose(award_price_mapping.get(x['IPN'], np.nan), x['PS Award Price'],
                                            atol=1e-5) else 'N', axis=1
            )
            # ------------------ end of Price match between CPN and awards dataframe -----------------------------

            # ------------------ UPDATE AWARDS DETAILS IN ACTIVE SUPPLIER DATAFRAME ------------------

            # Find matching rows in 'Awards' where 'Award CPN' matches the 'ipn' (Item Part Number)
            matching_indices = awards_df['Award CPN'] == ipn

            # Check if there are any matching indices
            if matching_indices.any():

                # Convert 'End Date' to datetime and then to the desired string format in 'awards_df'
                # We are modifying 'awards_df' in place
                awards_df.loc[matching_indices, 'End Date'] = pd.to_datetime(
                    awards_df.loc[matching_indices, 'End Date'],
                    errors='coerce').dt.strftime('%m-%d-%Y')

                # Drop rows where 'End Date' conversion resulted in NaT (not a time) to ensure we have valid end dates
                valid_end_dates = awards_df.loc[matching_indices].dropna(subset=['End Date'])

                # Check if there are any rows with valid end dates
                if not valid_end_dates.empty:

                    # Find the latest 'End Date' among the valid ones
                    latest_end_date = valid_end_dates['End Date'].max()

                    # Update 'PS Award Exp Date' in 'active_supplier_df' with the latest 'End Date'
                    # We are bringing in 'End Date' from 'awards_df' to 'active_supplier_df'
                    active_supplier_df.at[idx, 'PS Award Exp Date'] = latest_end_date

                    # Update 'PS Award Price' in 'active_supplier_df' if available
                    # We are bringing in 'Award Price' from 'awards_df' to 'active_supplier_df'
                    # This assumes the 'Award Price' is associated with the latest 'End Date'
                    if pd.notna(valid_end_dates['Award Price'].iloc[0]):
                        active_supplier_df.at[idx, 'PS Award Price'] = valid_end_dates['Award Price'].iloc[0]

                    # Update 'PS Awd Cust ID' in 'active_supplier_df' if available
                    # We are bringing in 'Award Cust ID' from 'awards_df' to 'active_supplier_df'
                    # This assumes the 'Award Cust ID' is associated with the latest 'End Date'
                    if pd.notna(valid_end_dates['Award Cust ID'].iloc[0]):
                        active_supplier_df.at[idx, 'PS Awd Cust ID'] = valid_end_dates['Award Cust ID'].iloc[0]

            # ----------------------- End of Update Awards Detail in active dataframe -----------------------

        # Convert the 'PS Award Exp Date' in active_supplier_df to 'MM-DD-YYYY' format
        active_supplier_df['PS Award Exp Date'] = pd.to_datetime(active_supplier_df['PS Award Exp Date'],
                                                                 errors='coerce').dt.strftime('%m-%d-%Y')
        # --------------------- End of: UPDATE AWARDS DETAILS IN ACTIVE SUPPLIER DATAFRAME-----------

        # ---------------------- Contract Change Logic -----------------------------------------------
        active_supplier_df['Contract Change'] = np.where(
            active_supplier_df['LW PRICE'].isna(),  # Check if 'LW PRICE' is NaN or null
            'New Item',  # New Item is going to be populated if the item was not in the prev contract
            np.where(  # Locate where they are going to include the other contingencies
                active_supplier_df['Price'] == active_supplier_df['LW PRICE'],
                'No Change',
                np.where(
                    active_supplier_df['Price'] > active_supplier_df['LW PRICE'],
                    'Price Increase',
                    'Price Decrease'  # Since we've covered all other scenarios, this can be the else condition.
                )
            )
        )
        # --------------------- End of Contract Change Logic --------------------------------------------

        # --------------------- Save Output File Logic --------------------------------------------------
        # Ask the user for the output file path
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save the output file as",
                                                   initialdir="P:\Partnership_Python_Projects\Creation\test_001")
        # --------------------- End of Save Output File Logic -------------------------------------------

        # --------------------- Saving the desired sheets in final output ------------------------------
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
                # --------------------- End of Saving the desired sheets in final output -----------------

                # --------------------- Additional formatting for specific columns -----------------------------
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

            process = psutil.Process()
            print(f"Memory usage: {process.memory_info().rss / 1024 ** 2:.2f} MB")  # memory usage in MB

    except Exception as e:
        messagebox.showerror("Error Process was Cancelled", str(e))
        process = psutil.Process()
        print(f"Memory usage at error: {process.memory_info().rss / 1024 ** 2:.2f} MB")
        # Exception is to cause the button to clear ability to re-click once complete
        # Adjusting the rate of how the file is being used and how much memory we are using
