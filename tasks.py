from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive

import shutil


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
        slowmo=200,
    )
    open_robot_order_website()
    get_orders()
    fill_from_csv()
    archive_receipts()
    clean_up()


def open_robot_order_website():
    '''
    Opens robot order website (no shit) | 
    Gives up constitutional rights
    '''
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click("button:text('Yep')")

def get_orders():
    """
    Downloads csv file from the given URL
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def fill_the_robot_order(order):
    """Fills in the robot order as per the customer order and clicks the 'Order' button"""
    page = browser.page()
    headnum = order["Head"]
    page.locator("#head").select_option(headnum)
    bodynum = "#id-body-" + order["Body"]
    page.locator(bodynum).click()
    page.fill("input[placeholder='Enter the part number for the legs']", order["Legs"])
    page.fill("#address", order["Address"])
    while True:
        page.click("#order")
        order_another = page.query_selector("#order-another")
        if order_another:
            pdf_path = store_receipt_as_pdf(int(order["Order number"]))
            screenshot_path = screenshot_robot(int(order["Order number"]))
            embed_screenshot_to_receipt(screenshot_path, pdf_path)
            order_another_bot()
            close_popup()
            break



def fill_from_csv():
    '''Fills out the form according to table data'''
    library = Tables()
    robot_orders = library.read_table_from_csv("orders.csv")
    for order in robot_orders:
        fill_the_robot_order(order)

def store_receipt_as_pdf(order_number):
    """This stores the robot order receipt as pdf"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = "output/receipts/{0}.pdf".format(order_number)
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """Take a screenshot of the ordered robot"""
    page = browser.page()
    screenshotpath = "output/screenshots/{0}.png".format(order_number)
    page.locator("#robot-preview-image").screenshot(path=screenshotpath)
    return screenshotpath

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """appends screenshot of robot to the pdf"""
    pdf = PDF()

    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=pdf_file
    )

def order_another_bot():
    """clicks on the order another button"""
    page = browser.page()
    page.locator("#order-another").click()

def close_popup():
    """closes popup for future robot orders"""
    page = browser.page()
    page.click("button:text('Yep')")

def archive_receipts():
    """Archives all receipts into one zip file"""
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")

def clean_up():
    """Cleans up the folders where receipts and screenshots are saved."""
    shutil.rmtree("./output/receipts")
    shutil.rmtree("./output/screenshots")