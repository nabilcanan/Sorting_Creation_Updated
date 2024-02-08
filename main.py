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
# from add_running import add_running_file_to_workbook
import webbrowser
import os

# import pygame
# from threading import Thread


# def play_background_music():
#     pygame.mixer.init()
#     pygame.mixer.music.load('images-videos/restaurant-music-110483.mp3')
#     pygame.mixer.music.play(-1)  # Play the music, -1 means play indefinitely in loop
#     pass


warnings.simplefilter('ignore', UserWarning)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def open_powerpoint():
    powerpoint_file = r'"P:\Partnership_Python_Projects\Creation\Creation_Python_Program.pptx"'

    try:
        os.startfile(powerpoint_file)
    except Exception as e:
        print(f"Error: {e}")


def open_readme_link():
    webbrowser.open('https://github.com/nabilcanan/Sorting_Creation_Updated/blob/main/README.md', new=2)


class ExcelSorter:
    def __init__(self):
        self.sort_last_ship_date_button = None
        self.sort_backlog_button = None
        self.sort_award_button = None
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

        # music_thread = Thread(target=play_background_music, daemon=True)
        # music_thread.start()

    def create_widgets(self, frame):
        style = ttk.Style()
        style.configure("TButton", font=("Rupee", 16, "bold"), width=60, height=2)
        style.map("TButton",
                  foreground=[('active', 'red')],
                  background=[('active', 'blue')])
        style.configure("TButton", background="white")  # Change the button background color to white

        title_label = ttk.Label(frame, text="Welcome Partnership Team!",
                                font=("Rupee", 32, "underline"), background="white", foreground="#103d81")
        title_label.pack(pady=10)

        open_powerpoint_button = ttk.Button(self.inner_frame, text='Open PowerPoint Instructions',
                                            command=open_powerpoint)
        open_powerpoint_button.pack(pady=10)

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

        print("Run Queries Button Called")
        new_instructions = ttk.Label(
            frame,
            text="For the Sort Query Files Button you will select all\n of your Raw Queries at once for sorting.",
            font=("Rupee", 19),
            background="white",
            anchor="center",
            justify="center",
            wraplength=1000
        )
        new_instructions.pack(pady=10)

        sort_multiple_files_button = ttk.Button(frame, text="Sort Query Files", command=self.sort_multiple_files,
                                                style="TButton")
        sort_multiple_files_button.pack(pady=10)
        print("Sort Multiple Files Button Called")

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

        merge_and_create_lost_items_button = ttk.Button(frame,
                                                        text="Merge Files and Create 'Lost Items' Sheet",
                                                        command=lambda: merge_files_and_create_lost_items(
                                                            merge_and_create_lost_items_button),
                                                        style="TButton")
        merge_and_create_lost_items_button.pack(pady=10)
        print("Merge Button Called")

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
                                            command=lambda: perform_vlookup(perform_vlookup_button), style="TButton")
        perform_vlookup_button.pack(pady=10)
        print("Perform V-lookup Button Called")

        logo_label = ttk.Label(frame, background="white")
        logo_label.pack(pady=10)

        logo_image = Image.open('images-videos/Sager-logo.png')
        logo_image = ImageTk.PhotoImage(logo_image)
        logo_label.config(image=logo_image)
        logo_label.image = logo_image

        # Center all the widgets vertically in the frame
        for widget in frame.winfo_children():
            widget.pack_configure(pady=5)

    def sort_multiple_files(self):
        print("Sort Multiple Files called")
        file_paths = filedialog.askopenfilenames(title="Select Query Excel files to sort",
                                                 filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")),
                                                 initialdir="P:\Partnership_Python_Projects\Creation\python_creation_setup_demo")
        if file_paths:
            for file_path in file_paths:
                if "Awards" in file_path:
                    print(f"Sorting '{file_path}' as an Awards file.")
                    self.sort_excel(file_path, ['Product ID', 'Award Cust ID'], [True, False], "Award")
                elif "Backlog" in file_path:
                    print(f"Sorting '{file_path}' as a Backlog file.")
                    self.sort_excel(file_path, ['Product ID', 'Backlog Entry'], [True, False], "Backlog")
                elif "Sales" in file_path:
                    print(f"Sorting '{file_path}' as a Sales History file.")
                    self.sort_excel(file_path, ['Product ID', 'Last Ship Date'], [True, False], "Sales History")
                elif "SND" in file_path:
                    print(f"Sorting '{file_path}' as a SND file.")
                    self.sort_excel(file_path, ['Product ID', 'SND Cost'], [True, True], "Ship & Debit")
                elif "VPC" in file_path:
                    print(f"Sorting '{file_path}' as a VPC file.")
                    self.sort_excel(file_path, ['PART ID', 'VPC Cost'], [True, False], "VPC")
                else:
                    print(f"File type for '{file_path}' not recognized. Skipping.")

            # After sorting all files, display a success message
            messagebox.showinfo("Success", "All your queries have been sorted.")
            print("All selected files have been sorted successfully.")
        else:
            # This message will be shown if no files were selected
            messagebox.showinfo("No Files Selected", "No files were selected for sorting.")
            print("No files were selected for sorting.")

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

    @staticmethod  # Write to data sheet for dataframes that are finished and merged
    def write_data_to_sheet(sheet, df):
        for r in dataframe_to_rows(df, index=False, header=True):
            sheet.append(r)

    def run(self):
        self.window.mainloop()


# Create an instance of the ExcelSorter and run the program
sorter = ExcelSorter()
sorter.run()
