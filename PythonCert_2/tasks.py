from robocorp.tasks import  task
from robocorp   import  browser
from RPA.HTTP import HTTP
from RPA.PDF import PDF
import pandas   as  pd
from zipfile import ZipFile
import time
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the rbot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
                slowmo=100,
        )
    
    open_robot_order_website()
    close_annoying_modal()
    get_orders()
    submit_orders()
    compress_directory('output/Receipts', 'output/Receipts.zip')

def open_robot_order_website():
    page = browser.page()

    page.goto("https://robotsparebinindustries.com/#/robot-order")

def read_csv_to_df(file_path):
    """Read the CSV file into a DataFrame"""
    df = pd.read_csv(file_path)
    return df

def screenshot_robot(image_id):
        """Take a screenshot of the page displying the sales results"""
        page = browser.page()
        robot_preview_image_locator = page.locator("#robot-preview-image")
        page.wait_for_selector("#robot-preview-image")
        page.screenshot(path=f"output/robotImages/robot_{image_id}.png")
        #page.screenshot(path=f"output/robotImages/robot_{image_id}.png", mask=[robot_preview_image_locator])
        return  f"output/robotImages/robot_{image_id}.png"

def embed_screenshot_to_pdf(screenshot_path, pdf_path):
    """Ensure that the directory exists"""
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    """Create a new pdf file and imbed the screenshot"""
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[screenshot_path],
        target_document=pdf_path
    )

def collect_results(robot_id):
    screenshot_path = screenshot_robot(robot_id)
    embed_screenshot_to_pdf(screenshot_path, f"output/Receipts/robot_{robot_id}.pdf")

def compress_directory(directory_path, output_zip_file):
    with ZipFile(output_zip_file, 'w') as zipf:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                if file.endswith('.pdf'):
                    # Create a complete file path
                    file_path = os.path.join(root, file)
                    # Add file to the zip file
                    zipf.write(file_path, os.path.relpath(file_path, directory_path))

def delete_orders_csv_file(file_path):
    """Tries tp delete the file at the given file path"""
    try:
        os.remove(file_path)
        print(f"File {file_path} has been deleted successfully.")
    except FileNotFoundError:
        print(f"The file {file_path} does not exist.")
    except PermissionError:
        print(f"Permission denied: unable to delete {file_path}.")
    except Exception as e:
        print(f"An error occurred: {e}.")

def get_orders():
    """Downloads excel file from the give URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", target_file=r'DataSets\orders.csv', overwrite=True)
    orders_df = read_csv_to_df('DataSets\orders.csv')
    delete_orders_csv_file('DataSets\orders.csv')
    return orders_df
        
def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")

def check_for_server_error():
    page = browser.page()

    error_selector = "div.alert.alert-danger[role='alert']"
    while   page.locator(error_selector).count() > 0:
        print("Server error message detected, proceeding to click 'Order' again")
        page.wait_for_selector("#order")
        page.click("#order")

def fill_the_form(row):
    page = browser.page()

    page.wait_for_selector("#head")
    page.select_option("#head", str(row['Head']))

    if  str(row["Body"]) == "1":
        page.wait_for_selector("#id-body-1")
        page.click("#id-body-1")

    if  str(row["Body"]) == "2":
        page.wait_for_selector("#id-body-2")
        page.click("#id-body-2")

    if  str(row["Body"]) == "3":
        page.wait_for_selector("#id-body-3")
        page.click("#id-body-3")

    if  str(row["Body"]) == "4":
        page.wait_for_selector("#id-body-4")
        page.click("#id-body-4")

    if  str(row["Body"]) == "5":
        page.wait_for_selector("#id-body-5")
        page.click("#id-body-5")

    if  str(row["Body"]) == "6":
        page.wait_for_selector("#id-body-6")
        page.click("#id-body-6")

    page.wait_for_selector('input[placeholder="Enter the part number for the legs"]')
    page.fill('input[placeholder="Enter the part number for the legs"]', str(row['Legs']))

    page.wait_for_selector("#address")
    page.fill("#address", row['Address'])

    page.wait_for_selector("#preview")
    page.click("#preview")

    collect_results(str(row['Order number']))
    
    page.wait_for_selector("#order")
    page.click("#order")

    check_for_server_error()
    
    page.wait_for_selector("#order-another")
    page.click("#order-another")

def submit_orders():
    """Gets order data, loops through the orders table and submits each order in the web form"""
    orders_df = get_orders()
    print(orders_df)

    for index, row in orders_df.iterrows():
        print(index, row['Order number'])
        fill_the_form(row)
        close_annoying_modal()

    