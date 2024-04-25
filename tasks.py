from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    download_csv_file()
    read_csv_file_and_complete_orders_and_save_receipts()
    archive_receipts()

def open_robot_order_website():
    """Opens the robot order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads cvs file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_csv_file_and_complete_orders_and_save_receipts():
    """Reads the downloaded cvs file into a table, uses that table to complete all orders, saving receipts with robot images along the way"""
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        page = browser.page()
        page.select_option("#head", order["Head"])
        if order["Body"] == "1":
            page.click("#id-body-1")
        elif order["Body"] == "2":
            page.click("#id-body-2")
        elif order["Body"] == "3":
            page.click("#id-body-3")
        elif order["Body"] == "4":
            page.click("#id-body-4")
        elif order["Body"] == "5":
            page.click("#id-body-5")
        elif order["Body"] == "6":
            page.click("#id-body-6")
        page.fill("input[type='number']", order["Legs"])
        page.fill("#address", order["Address"])
        page.click("#preview")
        page = browser.page()
        page.click("#order")
        page = browser.page()
        while page.is_hidden("#order-another"):
            page.click("#order")
            page = browser.page()
        receipt_html = page.locator("#receipt").inner_html()
        pdf = PDF()
        pdf.html_to_pdf(receipt_html, "receipt_text_temp.pdf")
        robot_image = page.locator("#robot-preview-image")
        robot_image.screenshot(path="robot_image_temp.png")
        order_number = order["Order number"]
        list_of_files = ["receipt_text_temp.pdf", "robot_image_temp.png"]
        pdf.add_files_to_pdf(files=list_of_files, target_document=f"receipts/receipt{order_number}.pdf", append=False)
        page.click("#order-another")
        

def get_orders():
    """Returns the order data as a table"""
    library = Tables()
    result = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return result

def close_annoying_modal():
    """Closes the pop-up on the robor order website"""
    page = browser.page()
    page.click("text=Yep")

def archive_receipts():
    """Adds the receipts to a zip file"""
    lib = Archive()
    lib.archive_folder_with_zip("receipts", "output/receipts.zip")