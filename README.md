# Partnership Sort Creation

The "Partnership Sort Creation" program is designed to sort the creation files for the week and perform various data operations within them. This program helps organize and consolidate the data for efficient processing.

## Requirements

- `pandas`
- `openpyxl`
- `Python 3.11`
- `Pycharm`
- `PIL`
- `Latest version of Python`


## Functionality

1. Sorting Creation Files:
   - The program allows you to sort different types of creation files, such as Award, Backlog, Sales History, SND, and VPC files.
   - Sorting is performed based on columns we need sorted
   - The sorted files are saved within the same file.

2. Adding Files to Active Contract File:
   - The program facilitates merging multiple files into the Active Contract File.
   - The files include the Active Supplier Contracts, Prev Contract, Awards, Backlog, Sales History, SND, and VPC files.
   - The program adds these files as separate sheets in the Active Contract File.
   - Additionally, it identifies and creates a "Lost Items" sheet that shows missing items from the previous to the active contract.

3. Updating Active Contract File with Previous Contract Information:
   - The program retrieves the previous contract file and copies the data into a new sheet named "Prev Contract" in the Active Contract File.
   - It then performs data operations to match the IPNs from the previous and active contracts.
   - The program adds the columns from the previous contract to the Active Supplier Contracts sheet, filling in the corresponding data for matching IPNs.
   - The merged data is saved in the Active Contract File.

## How to Use

1. Run the program and select the desired file to perform sorting or merging operations.
2. Follow the prompts and select the appropriate options to sort or merge the files.
3. The program will automatically perform the operations and save the resulting files.



Please ensure that you have the necessary dependencies installed before running the program. Refer to the "Requirements" section for details.

For any issues or questions, please contact Nabil Canan.

![Logo](images/Sager-logo.png)
