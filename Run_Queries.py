from pywinauto import ElementNotFoundError
from pywinauto.application import Application
import time
import pyautogui
import os
import warnings

warnings.simplefilter("ignore")


class Run_Queries:
    @staticmethod
    def click_button_image(image_path, confidence=0.8):
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            center = pyautogui.center(location)
            pyautogui.click(center)
        except TypeError:
            print(f"Image '{image_path}' not found on screen!")

    @staticmethod
    def run_query_and_export(query_name, window_title, image_path, next_query_name):
        try:
            app = Application().start(r'C:\FS760\bin\CLIENT\WINX86\pstools.exe')

            signon_window = app['PeopleSoft Signon']
            signon_window.wait('ready', timeout=20)
            print("'PeopleSoft Signon' window is ready.")

            username_field = signon_window.child_window(class_name="Edit", found_index=1)
            username_field.set_focus().type_keys('NCANAN', with_spaces=True)
            print("Username entered successfully.")

            password_field = signon_window.child_window(class_name="Edit", found_index=2)
            password_field.set_focus().type_keys('Jesus9637ever', with_spaces=True)
            print("Password entered successfully.")

            signon_window.child_window(title="OK", class_name="Button").click()
            print("OK button clicked.")
            time.sleep(3)

            app = Application().connect(title_re="Application Designer - .*")
            app.top_window().menu_select("Go->PeopleTools->Query")
            app.top_window().close()
            time.sleep(2)

            query_app = Application().connect(title="Untitled - Query")
            print("'Untitled - Query' window connected.")

            toolbar = query_app['Untitled - Query'].child_window(class_name="ToolbarWindow32")
            toolbar.button(1).click_input()
            print("'Open Query' icon clicked.")
            time.sleep(1)

            open_query_app = Application().connect(title="Open Query")
            open_query_window = open_query_app['Open Query']
            print("'Open Query' window connected.")

            open_query_window.Edit.set_text(query_name)
            print(f"Typed query name: {query_name}.")

            open_query_window.OK.click_input()
            print(f"Pressed OK button to search for query.")
            time.sleep(1)

            new_window_title = window_title
            new_query_app = Application().connect(title_re=new_window_title)
            new_query_window = new_query_app.window(title_re=new_window_title)
            new_query_window.wait('visible', timeout=5)
            print(f"Connected to new window: {new_window_title}")

            time.sleep(2)

            if new_query_window.exists(timeout=20):
                Run_Queries.click_button_image(image_path)
                print("Clicked the Run to Excel button using pyautogui.")

                time.sleep(20)

                excel_app = Application().connect(class_name="XLMAIN")
                excel_window = excel_app.window(class_name="XLMAIN")
                excel_window.wait('active', timeout=15)
                print("Excel window is active.")
                time.sleep(5)

                toolbar.button(1).click_input()
                time.sleep(5)
                open_query_window.Edit.set_text(next_query_name)
                open_query_window.OK.click_input()

                print(f"Opened the next query: {next_query_name}")

            else:
                print(f"Window '{new_window_title}' not found or already closed.")

            if new_query_window.exists():
                new_query_window.close()
                print(f"Closed query window: {window_title}")

        except ElementNotFoundError as e:
            print(f"Element not found: {str(e)}")
        except Exception as e:
            print(f"Error: {str(e)}")

    queries = [
        {"query_name": "PUBLIC.QUERY.STRATEGIC_ACTIVE_AWARDS.ALL ACTIVE AWARD BY GROUP",
         "window_title": "PUBLIC.QUERY.STRATEGIC_ACTIVE_AWARDS - ALL ACTIVE AWARD BY GROUP - Query",
         "next_query_name": "Next Query 1 Name"},
        {"query_name": "PUBLIC.QUERY.STRATEGIC_OPEN_ORDERS.OOR BY STRATEGIC GROUP-7-20-18",
         "window_title": "PUBLIC.QUERY.STRATEGIC_OPEN_ORDERS - OOR by Strategic Group-7-20-18 - Query",
         "next_query_name": "Next Query 2 Name"},
        {"query_name": "PUBLIC.QUERY.STRATEGIC_SOLDTOS_SND.STRATEGIC SOLDTOS SND",
         "window_title": "PUBLIC.QUERY.STRATEGIC_SOLDTOS_SND - Strategic Soldtos SND - Query",
         "next_query_name": "Next Query 3 Name"},
        {"query_name": "PUBLIC.QUERY.STRATEGIC_ACTIVE_VPC.ALL ACTIVE VPCS FOR STRATEGIC",
         "window_title": "PUBLIC.QUERY.STRATEGIC_ACTIVE_VPC - all active VPCs for strategic - Query",
         "next_query_name": "Next Query 4 Name"},
        {"query_name": "PUBLIC.QUERY.STRATEGIC_SALES.SALES BY CUSTOMER GROUP",
         "window_title": "PUBLIC.QUERY.STRATEGIC_SALES - Sales by Customer Group - Query",
         "next_query_name": "Next Query 5 Name"}
    ]

    # Path to the "Run to Excel" button image
    current_directory = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(current_directory, 'images-videos', 'run_to_excel.png')

    def run_all_queries(self):
        try:
            for query in self.queries:
                query_name = query["query_name"]
                window_title = query["window_title"]
                next_query_name = query["next_query_name"]
                self.run_query_and_export(query_name, window_title, self.image_path, next_query_name)

            print("Process completed successfully!")

        except Exception as e:
            print(f"Error: {str(e)}")


if __name__ == "__main__":
    runner = Run_Queries()
    runner.run_all_queries()
