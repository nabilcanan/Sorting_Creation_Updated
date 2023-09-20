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
_______________________________________________________________________________________________________________

# Functionality For Buttons
_______________________________________________________________________________________________________________


- For the **Run Queries Button** please enter your Peoplesoft Credentials & do NOT move your mouse until all the queries have been run

_______________________________________________________________________________________________________________

- For the **Sort Awards Button** please select the awards file you have from your Run Queries
- For the **Sort Backlog Button** please select the backlog file you have from your Run Queries
- For the **Sort Sales History Button** please select the sales history file you have from your Run Queries
- For the **Sort SND File Button** please select the SND file you have from your Run Queries
- For the **Sort VPC Button** please select the VPC file you have from your Run Queries

_______________________________________________________________________________________________________________

- For the **Merge Files and Create** 'Lost Items' Sheet button** select the files in this order:

1. Current Weeks Creation Contract file
2. Last Week's Creation Contract File
3. Sorted Awards File
4. Sorted Backlog File
5. Sorted Sales History File
6. Sorted SND File
7. Sorted VPC File
8. Latest 30 Day Running File
_______________________________________________________________________________________________________________

- For the **Perform VLook-Up** button select in this order 

1. All you need to select for this button is the new workbook we just created with the lost items sheet, and where we have all of our combined sorted sheets. 
2. Once the process is complete you will be prompt to save your new file.

_______________________________________________________________________________________________________________

- For the **Add Latest Running File** button select the files in this order

1. The File where we need to add the Running File (In your case it will be the file where we have our lost items, the sorted sheets, and where the VLOOKUP was completed.)
2. The Latest Running 30 Day File that can be found in our Creation folder
_______________________________________________________________________________________________________________

There is a video on this program working step by step in the images-video folder, to access just clock the folder in the repository, then clock the video (mkv file) and select 'view raw'
This will download the video for you to view.

Please ensure that you have the necessary dependencies installed before running the program. Refer to the "Requirements" section for details.

For any issues or questions, please contact Nabil Canan.

![Logo](images-videos/Sager-logo.png)
