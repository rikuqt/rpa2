from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
import zipfile
import glob


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    open_order_website()
    download_csv_file()
    read_the_orders()


def open_order_website():
    """Navigates to the order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def download_csv_file():
    """Downloads the CSV file with the orders"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def read_the_orders():
    """Reads the orders from the CSV file"""
    tables = Tables()
    tables_data = tables.read_table_from_csv("orders.csv")
 

    for order in tables_data:
        fill_order_form(order)


def fill_order_form(order):
    """Fills in the order form"""
    

    close_annoying_modal()
    page = browser.page()

    page.select_option("#head", str(order["Head"]))
    page.check(f"input[type='radio'][value='{order['Body']}']")
    page.fill("input[placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", order["Address"])

    while page.query_selector("id=order") != None:
        try: 
            robot_preview()
            page.click("button:has-text('Order')")
        except Exception:
            pass

    
    order_number = page.text_content("p.badge.badge-success")



    pdf = store_receipt_as_pdf(order_number)
    screenshot = take_screenshot(order_number)

    embed_screenshot_to_receipt(pdf, screenshot)
    archive_receipts()
    order_another()

    

def order_another():
    """Orders another robot"""
    page = browser.page()
    page.click("button:has-text('Order another robot')")

def robot_preview():
    """Previews the robot"""
    page = browser.page()
    page.click("button:has-text('Preview')")



def store_receipt_as_pdf(order_number):
    """Stores the receipt as a PDF file"""
    pdf = PDF()
    page = browser.page()
    robot_receipt = page.locator(selector="#receipt").inner_html()

    
    pdf_path = f"output/{order_number}.pdf"
    pdf.html_to_pdf(robot_receipt, pdf_path)
    return pdf_path
    

def take_screenshot(order_number):
    """Takes screenshots of the ordered robot and another element"""
    page = browser.page()
    screenshot_path = f"output/{order_number}.png"
    
    # Take screenshot of the robot preview image
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
   
    return screenshot_path


def embed_screenshot_to_receipt(pdf_file, screenshot):
    """Embeds the screenshot to the PDF receipt"""
    pdf = PDF()

    list_of_files = [pdf_file, screenshot]
    pdf.add_files_to_pdf(
        list_of_files,
        pdf_file
    )


def archive_receipts():
    """Creates ZIP archive of the receipts and the images"""
    with zipfile.ZipFile("output/robot_orders.zip", 'w') as zipf:
        for file in glob.glob("output/*.pdf") + glob.glob("output/*.png"):
            zipf.write(file, arcname=file.split('/')[-1])


def close_annoying_modal():
    """Closes the annoying modal"""
    page = browser.page()
    page.click("text=Yep")


