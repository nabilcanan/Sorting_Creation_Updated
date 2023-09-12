import time
from tkinter import messagebox
from pywinauto import ElementNotFoundError
from pywinauto.application import Application
import pyautogui
import os


def click_button_image(image_path, confidence=0.8, offset=0, double_click_required=False):
    try:
        print(f"Looking for image '{image_path}' on screen...")
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)
        center = pyautogui.center(location)

        new_click_location = (center[0] + offset, center[1])
        time.sleep(3)  # wait a bit before clicking

        if 'WHERETOCLICKIMG4' in image_path:
            pyautogui.doubleClick(x=1096, y=523)  # Specific coordinates for WHERETOCLICKIMG4
        elif double_click_required:
            pyautogui.doubleClick(new_click_location)
        else:
            pyautogui.click(new_click_location)  # For other images, single click

        print(f"Successfully clicked on '{image_path}'.")
        time.sleep(3)  # wait a bit after clicking

    except TypeError:
        print(f"Image '{image_path}' not found on screen!")


def new_function():
    image_path_run_to_excel = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'run_to_excel.png')
    image_path_criteria = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CRITERIAPANEL.PNG')
    image_path_click1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WHERETOCLICKIMG1.png')
    image_path_click2 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WHERETOCLICKIMG2.png')
    image_path_click3 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WHERETOCLICKIMG3.png')
    image_path_click4 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WHERETOCLICKIMG4.png')
    image_path_click5 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WHERETOCLICKIMG5.png')
    image_path_no_button = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'no_image_click.PNG')

    query1 = "PUBLIC.QUERY.STRATEGIC_ACTIVE_AWARDS.ALL ACTIVE AWARD BY GROUP"
    query2 = "PUBLIC.QUERY.STRATEGIC_OPEN_ORDERS.OOR BY STRATEGIC GROUP-7-20-18"
    query3 = "PUBLIC.QUERY.STRATEGIC_ACTIVE_VPC.ALL ACTIVE VPCS FOR STRATEGIC"
    query4 = "PUBLIC.QUERY.STRATEGIC_SOLDTOS_SND.STRATEGIC SOLDTOS SND"
    query5 = "PUBLIC.QUERY.STRATEGIC_SALES.SALES BY CUSTOMER GROUP"

    def handle_save_changes_prompt():
        try:
            print("Checking for Save Changes prompt...")
            time.sleep(2)
            click_button_image(image_path_no_button)
            print("Clicked 'No' button.")
            time.sleep(1)
        except TypeError:
            print("Save Changes prompt not found. Continuing...")

    def login_and_run_query(query, where_to_click_image):
        print(f"Starting '{query}'...")

        # Get user credentials
        username = input("Please enter your username: ")
        password = input("Please enter your password: ")

        # Launch and Connect to PeopleSoft
        app = Application().start(r'C:\FS760\bin\CLIENT\WINX86\pstools.exe')
        signon_window = app['PeopleSoft Signon']
        signon_window.wait('ready', timeout=20)

        # Enter credentials and Login
        username_field = signon_window.child_window(class_name="Edit", found_index=1)
        username_field.set_focus().type_keys(username, with_spaces=True)

        password_field = signon_window.child_window(class_name="Edit", found_index=2)
        password_field.set_focus().type_keys(password, with_spaces=True)

        signon_window.child_window(title="OK", class_name="Button").click()
        time.sleep(8)

        # WE CAN USE THIS CODE SEGMENT ^^^^^ OR THE BOTTOM ONE WHERE WE AUTO ENTER THE USERS INFO
        # # Launch and Connect to PeopleSoft
        # app = Application().start(r'C:\FS760\bin\CLIENT\WINX86\pstools.exe')
        # signon_window = app['PeopleSoft Signon']
        # signon_window.wait('ready', timeout=20)
        #
        # # Enter credentials and Login
        # username_field = signon_window.child_window(class_name="Edit", found_index=1)
        # username_field.set_focus().type_keys('NCANAN', with_spaces=True)
        #
        # password_field = signon_window.child_window(class_name="Edit", found_index=2)
        # password_field.set_focus().type_keys('Jesus9637ever', with_spaces=True)
        #
        # signon_window.child_window(title="OK", class_name="Button").click()
        # time.sleep(8)

        # Go to Query menu
        app = Application().connect(title_re="Application Designer - .*")
        app.top_window().menu_select("Go->PeopleTools->Query")
        app.top_window().close()
        time.sleep(5)

        # Open Query
        query_app = Application().connect(title="Untitled - Query")  # Modified line
        toolbar = query_app.top_window().child_window(class_name="ToolbarWindow32")
        toolbar.button(1).click_input()
        time.sleep(5)

        open_query_app = Application().connect(title="Open Query")
        open_query_window = open_query_app['Open Query']

        open_query_window.Edit.set_text(query)
        print(f"Query text for '{query}' set. Waiting a bit before clicking OK...")
        time.sleep(3)
        open_query_window.OK.click_input()
        print(f"OK clicked for '{query}'. Waiting for criteria panel...")
        time.sleep(5)

        # Click on criteria panel
        click_button_image(image_path_criteria)

        # Click on specified "where to click" image (either WHERETOCLICKIMG1.png or WHERETOCLICKIMG2.png)
        click_button_image(where_to_click_image, offset=50, double_click_required=True)  # <-- Updated line

        # Enter NEOTECH in the 'Constant' window and click OK
        constant_app = Application().connect(title="Constant")
        constant_app.window(title="Constant").child_window(class_name="Edit").set_text("CREATION")
        constant_app.window(title="Constant").child_window(title="OK", class_name="Button").click_input()
        # time.sleep(2)

        # Marker to check if run_to_excel has been clicked
        run_to_excel_clicked = False

        def click_run_to_excel():
            nonlocal run_to_excel_clicked
            if not run_to_excel_clicked:
                # Look for the image and click it
                click_button_image(image_path_run_to_excel)
                run_to_excel_clicked = True
                time.sleep(20)  # Wait for 20 seconds after clicking

        click_run_to_excel()

        try:
            if query_app.window(title_re=".*Query.*").exists():
                print("Closing query window...")
                query_app.window(title_re=".*Query.*").close()
                print("Query window closed.")
        except Exception as e:
            print(f"Error while closing the query window: {str(e)}")

    try:
        # Use a counter to keep track of completed queries
        queries_completed = 0

        def on_query_completed():
            nonlocal queries_completed
            queries_completed += 1
            print(f"Completed {queries_completed} queries.")
            if queries_completed == 5:
                messagebox.showinfo("Queries Completed",
                                    "Both queries have been executed. Please allow some time for Excel sheets to load.")

        # Run both queries
        print("Running first query...")
        login_and_run_query(query1, image_path_click1)
        handle_save_changes_prompt()
        print("Running second query...")
        login_and_run_query(query2, image_path_click2)
        handle_save_changes_prompt()
        print("Running third query...")
        login_and_run_query(query3, image_path_click3)
        handle_save_changes_prompt()
        print("Running fourth query...")
        login_and_run_query(query4, image_path_click4)
        handle_save_changes_prompt()
        print("Running fifth query...")
        login_and_run_query(query5, image_path_click5)
        handle_save_changes_prompt()

        on_query_completed()

    except ElementNotFoundError as e:
        print(f"Element not found: {str(e)}")
    except Exception as e:
        print(f"Error: {str(e)}") \
