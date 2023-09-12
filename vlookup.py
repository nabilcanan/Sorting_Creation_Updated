from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
import numpy as np


def perform_vlookup():
    print("perform_vlookup called")
    try:
        # Ask the user for the contract file paths
        contract_file = filedialog.askopenfilename(title="Select the contract file, where we need a vlookup",
                                                   initialdir="I:\Quotes\Partnership Sales - CM\Creation")

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
        messagebox.showerror("Error Process was Cancelled", str(e))
