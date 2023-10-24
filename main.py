import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
from PIL import ImageTk, Image
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import filedialog, messagebox
import warnings
from queries import new_function
from vlookup import perform_vlookup
from merge import merge_files_and_create_lost_items
from add_running import add_running_file_to_workbook
import webbrowser

warnings.simplefilter('ignore', UserWarning)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def open_readme_link():
    webbrowser.open('https://github.com/nabilcanan/Sorting_Creation_Updated/blob/main/README.md', new=2)


class ExcelSorter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Sorting Creation Files And Performing VLookUp")
        self.window.configure(bg="white")
        self.window.geometry("875x600")  # Usually 600 for normal window size

        # Create a canvas and a vertical scrollbar
        self.canvas = tk.Canvas(self.window)
        self.scrollbar = ttk.Scrollbar(self.window, orient="vertical", command=self.canvas.yview)

        # Configure the canvas to respond to the scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Create a frame to hold your widgets, and add it to the canvas
        self.inner_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((self.window.winfo_width() / 2, 0), window=self.inner_frame, anchor="n")

        # Configure the canvasAs's scroll-region to encompass the frame
        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Pack the scrollbar, making sure it sticks to the right side
        self.scrollbar.pack(side="right", fill="y")

        # Configure the canvas to expand and fill the window
        self.canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)

        # Canvas - Scrollbar
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.canvas.yview)

        self.create_widgets(self.inner_frame)
        style = ttk.Style()
        style.configure("TFrame", background="white")  # set the background color of ttk.Frame to white

        self.inner_frame = ttk.Frame(self.canvas, style="TFrame")  # create the inner frame with the updated style

        # This function will get triggered when the mouse wheel is scrolled
        def _on_mousewheel(event):
            self.canvas.yview_scroll(-1 * (event.delta // 120), "units")

        # Bind the function to the MouseWheel event, to make our scrolling function more applicable 
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def create_widgets(self, frame):
        style = ttk.Style()
        style.configure("TButton", font=("Rupee", 16, "bold"), width=60, height=2)
        style.map("TButton",
                  foreground=[('active', 'red')],
                  background=[('active', 'blue')])
        style.configure("TButton", background="white")  # Change the button background color to white

        title_label = ttk.Label(frame, text="Welcome Partnership Member!",
                                font=("Rupee", 32, "underline"), background="white", foreground="#103d81")
        title_label.pack(pady=10)

        description_label = ttk.Label(frame,
                                      text="This tool allows you to sort your Excel files for our Creation "
                                           "Contact",
                                      font=("Rupee", 20, "underline"), background="white")
        description_label.pack(pady=10)

        description_label = ttk.Label(frame, text="Select the Run Queries Button and don't move your mouse\n "
                                                  "    after you enter your Peoplesoft Login Credentials\n"
                                                  "    Make sure you notice the queries are run in this order\n"
                                                  "  1. Awards 2. Backlog 3. VPC 4. SND 5. Sales History",
                                      font=("Rupee",
                                            18, "bold"),
                                      background="white")
        description_label.pack(anchor='center')

        run_queries_button = ttk.Button(frame, text="Run Queries", command=new_function, style="TButton")
        run_queries_button.pack(pady=10)

        sort_award_button = ttk.Button(frame, text="Sort Award File", command=self.sort_award_file,
                                       style="TButton")
        sort_award_button.pack(pady=10)

        sort_backlog_button = ttk.Button(frame, text="Sort Backlog File", command=self.sort_backlog_file,
                                         style="TButton")
        sort_backlog_button.pack(pady=10)

        sort_last_ship_date_button = ttk.Button(frame, text="Sort Sales History File",
                                                command=self.sort_by_last_ship_date, style="TButton")
        sort_last_ship_date_button.pack(pady=10)

        sort_ship_and_debit = ttk.Button(frame, text="Sort SND File", command=self.sort_ship_and_debit,
                                         style="TButton")
        sort_ship_and_debit.pack(pady=10)

        sort_vpc = ttk.Button(frame, text="Sort VPC File", command=self.sort_vpc, style="TButton")
        sort_vpc.pack(pady=10)

        add_instructions_for_active_contracts_file = ttk.Label(
            frame,
            text="The 'Merge Files and Create 'Lost Items' Sheet' button allows you to\n"
                 "bring those sorted files all together accordingly into one workbook.\n"
                 "Order to Select Files:\n 1. Current Contract\n "
                 "2. Previous Weeks Contract\n 3. Awards File 4. Backlog File\n "
                 "5. Sales History File\n 6. SND File 7. VPC File\n"
                 "You will get a success message at the end",
            font=("Rupee", 19),
            background="white",
            anchor="center",
            justify="center",
            wraplength=1000
        )
        add_instructions_for_active_contracts_file.pack(pady=2)

        merge_and_create_lost_items_button = ttk.Button(frame, text="Merge Files and Create 'Lost Items' Sheet",
                                                        command=merge_files_and_create_lost_items, style="TButton")
        merge_and_create_lost_items_button.pack(pady=10)

        new_instructions = ttk.Label(
            frame,
            text="For the 'Perform VLOOKUP' button \n"
                 "1. Select the file where you now need your vlookup completed.\n"
                 "(This is the same file where all your files are now merged.) ",
            font=("Rupee", 19),
            background="white",
            anchor="center",
            justify="center",
            wraplength=1000
        )
        new_instructions.pack(pady=10)

        perform_vlookup_button = ttk.Button(frame, text="Perform VLook-Up to new file",
                                            command=perform_vlookup, style="TButton")
        perform_vlookup_button.pack(pady=10)

        add_instructions_for_active_contracts_file = ttk.Label(
            frame,
            text="For the 'Add Latest Running File' button\n"
                 "Select the Files in This order:\n"
                 "1. The file where we are adding the Running File.\n"
                 "2. The latest version of the Running file.",
            font=("Rupee", 19),
            background="white",
            anchor="center",
            justify="center",
            wraplength=1000
        )
        add_instructions_for_active_contracts_file.pack(pady=2)

        add_running_file_button = ttk.Button(frame, text="Add Latest Running File",
                                             command=add_running_file_to_workbook, style="TButton")
        add_running_file_button.pack(pady=10)

        description_label = ttk.Label(frame, text="Feel free to check out the ReadMe for more detailed instructions",
                                      font=("Rupee", 18),
                                      background="white")
        description_label.pack(anchor='center')

        readme_button = ttk.Button(frame, text="Open ReadMe",
                                   command=open_readme_link, style="TButton")
        readme_button.pack(pady=10)

        logo_label = ttk.Label(frame, background="white")
        logo_label.pack(pady=10)

        logo_image = Image.open('images-videos/Sager-logo.png')
        logo_image = ImageTk.PhotoImage(logo_image)
        logo_label.config(image=logo_image)
        logo_label.image = logo_image

        # Center all the widgets vertically in the frame
        for widget in frame.winfo_children():
            widget.pack_configure(pady=5)

    @staticmethod
    def select_file(file_type="Excel"):
        print("Select File Function  called")

        file_path = filedialog.askopenfilename(title=f"Select {file_type} file",
                                               filetypes=(
                                                   ("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")),
                                               initialdir="P:\Partnership_Python_Projects\Creation\test_001")

        if file_path:
            return file_path
        else:
            return None

    # Here is where we will sort all the files we ran our queries from
    def sort_award_file(self):
        print("Sort Award File called")
        file_path = self.select_file("Awards")
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'Award Cust ID'], [True, False], "Award")

    def sort_backlog_file(self):
        print("Sort Backlog File called")
        file_path = self.select_file("Backlog")
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'Backlog Entry'], [True, False], "Backlog")

    def sort_by_last_ship_date(self):
        print("Sort Sales File called")
        file_path = self.select_file("Sales")
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'Last Ship Date'], [True, False], "Sales History")

    def sort_ship_and_debit(self):
        print("Sort Ship and Debit called")
        file_path = self.select_file("SND")
        if file_path:
            self.sort_excel(file_path, ['Product ID', 'SND Cost'], [True, True], "Ship & Debit")

    def sort_vpc(self):
        print("Sort VPC File called")
        file_path = self.select_file("VPC")
        if file_path:
            self.sort_excel(file_path, ['PART ID', 'VPC Cost'], [True, False], "VPC")

    @staticmethod  # This is for some exceptions where we have certain numbers in our Excel file
    def sort_excel(file_path, sort_columns, ascending_order, file_type=""):
        if not sort_columns:
            messagebox.showerror("Error", "No columns selected for sorting.")
            return

        try:
            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(file_path)

            # If 'SND Cost' is one of the sort columns, convert it to numeric
            if 'SND Cost' in sort_columns:
                df['SND Cost'] = pd.to_numeric(df['SND Cost'], errors='coerce')

            # If 'VPC Cost' is one of the sort columns, convert it to numeric
            if 'VPC Cost' in sort_columns:
                df['VPC Cost'] = pd.to_numeric(df['VPC Cost'], errors='coerce')

            # Sort the DataFrame based on the selected columns
            df = df.sort_values(by=sort_columns, ascending=ascending_order)

            # Save the sorted DataFrame back to the Excel file
            df.to_excel(file_path, index=False)

            messagebox.showinfo("Success!", f"Success! {file_type} file sorted and saved successfully.")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    @staticmethod
    def write_data_to_sheet(sheet, df):
        for r in dataframe_to_rows(df, index=False, header=True):
            sheet.append(r)

    def run(self):
        self.window.mainloop()


# Create an instance of the ExcelSorter and run the program
sorter = ExcelSorter()
sorter.run()
