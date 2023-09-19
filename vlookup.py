from tkinter import filedialog, messagebox
import pandas as pd
# from openpyxl.styles import PatternFill
# from openpyxl.styles import Alignment
import numpy as np


def perform_vlookup():
    try:
        # Ask the user for the contract file paths
        contract_file = filedialog.askopenfilename(title="Select the contract file, where we need a vlookup",
                                                   initialdir="I:\Quotes\Partnership Sales - CM\Creation")

        # Load all the sheets
        active_supplier_df = pd.read_excel(contract_file, sheet_name='Active Supplier Contracts', header=1)
        prev_contract_df = pd.read_excel(contract_file, sheet_name='Prev Contract', header=0)
        lost_items_df = pd.read_excel(contract_file, sheet_name='Lost Items')
        awards_df = pd.read_excel(contract_file, sheet_name='Awards')
        snd_df = pd.read_excel(contract_file, sheet_name='SND')
        vpc_df = pd.read_excel(contract_file, sheet_name='VPC')
        backlog_df = pd.read_excel(contract_file, sheet_name='Backlog')
        sales_history_df = pd.read_excel(contract_file, sheet_name='Sales History')
        running_file_df = pd.read_excel(contract_file, sheet_name='Running File - 30 Day Notice Co')

        # Merge on 'IPN' to get the 'PSoft Part' column
        active_supplier_df = active_supplier_df.merge(prev_contract_df[['IPN', 'PSoft Part']], on='IPN', how='left')

        # Create and initialize the 'Cost' column
        active_supplier_df['Cost'] = np.nan

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

        # Ask the user for the output file path
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save the output file as")

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
                running_file_df.to_excel(writer, index=False, sheet_name='Running File - 30 Day Notice Co')

            # Display a success message in a message box
            messagebox.showinfo("Success", "The output file has been saved as: " + output_file)

    except Exception as e:
        messagebox.showerror("Error Process was Cancelled", str(e))
