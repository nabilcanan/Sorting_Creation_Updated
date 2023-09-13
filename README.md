# Partnership Sort Creation

The "Partnership Sort Creation" program is designed to sort the creation files for the week and perform various data operations within them. This program helps organize and consolidate the data for efficient processing.

## Requirements, these will need to be done by pip

- `pandas`
- `openpyxl`
- `Python 3.11`
- `Pycharm`
- `PIL`
- `PyWinAuto`
- `pyautogui`
- `os`
- `opencv-python`
- `image`


## Functionality

1. Sorting Creation Files(Sort Awards, Backlog, Sales History, etc):
   - The program allows you to sort different types of creation files, such as Award, Backlog, Sales History, SND, and VPC files.
   - Sorting is performed based on columns we need sorted that was in our Creation Instructions.
   - The sorted files are saved within the same file.

2. Adding Files to Active Contract File (Merge and Create 'Lost Items' Sheet):
   - The program facilitates merging multiple files into the Active Contract File.
   - The files include the Active Supplier Contracts, Prev Contract, Awards, Backlog, Sales History, SND, and VPC files.
   - The program adds these files as separate sheets in the Active Contract File.
   - Additionally, it identifies and creates a "Lost Items" sheet that shows missing items from the previous to the active contract.

3. Updating Active Contract File with VLOOK-UP Function:
   - Select the new file we created in our last button with all the files merged and your vlookup will occur
   - Once the process is complete you can choose where to save your new file
   
For the **Merge Files and Create 'Lost Items' Sheet button** select the files in this order:
1. current week
2. last week
3. awards
4. backlog
5. sales
6. snd
7. vpc
8. running file 

For the vlookup button select in this order 
1. All you need to select for this button is the new workbook we just created with the lost items sheet. 
2. Then choose where you'd like to save the final workbook with all your data. 

There is a video on this program working step by step in the images-video folder, to access just clock the folder in the repository, then clock the video (mkv file) and select 'view raw'
This will download the video for you to view.

Please ensure that you have the necessary dependencies installed before running the program. Refer to the "Requirements" section for details.

For any issues or questions, please contact Nabil Canan.

![Logo](images-videos/Sager-logo.png)
