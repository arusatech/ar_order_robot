from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Archive import Archive

import os


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)
    orders = get_orders()
    ar_print(orders)
    open_robot_order_website()
    order_multiple_robots(orders)
    archive_receipts()
    close_robot_order_website()

def ar_print(data, load=False, marshall=True, indent=2):
    def _stringify_val(data):
        if isinstance(data, dict):
            return {k: _stringify_val(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [_stringify_val(v) for v in data]
        elif isinstance(data, (str, int, float)):
            return data
        return str(data)

    _data = _stringify_val(data) if marshall else data
    try:
        _d = (
            json.dumps(json.loads(_data), indent=indent) if load else
            json.dumps(_data, indent=indent)
        )
    except:
        _d = _data

    print(_d)


def open_robot_order_website():
    """
    Opens the RobotSpareBin Industries Inc. website and handles the modal.
    """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    handle_modal()


def handle_modal():
    """
    Handles the modal by clicking the 'OK' button if it appears.
    """
    page = browser.page()
    try:
        page.click("button:text('OK')", timeout=5000)
    except:
        print("Modal not found or already closed.")


def get_orders():
    """
    Retrieves the orders from the CSV file.
    """
    http = HTTP()
    csv_content = http.http_get(url="https://robotsparebinindustries.com/orders.csv", stream=True).content.decode("utf-8")
    csv_lines = [line.split(',') for line in csv_content.splitlines()][1:]
    # ar_print(csv_lines)
    tables = Tables()
    return tables.create_table(csv_lines)

def order_multiple_robots(orders):
    """
    Orders multiple robots from RobotSpareBin Industries Inc.
    """
    page = browser.page()
    for order in orders:
        fill_order_form(page, order)
        submit_order(page)
        save_receipt_as_pdf(page, order)
        page.click("button:text('Order another robot')") # click the order more button   
        handle_modal()
        # Add code here to save the receipt and screenshot

def fill_order_form(page, order):
    """
    Fills out the order form for a single robot.
    """
    page.select_option("#head", str(order[1]))
    page.click(f"#id-body-{order[2]}")
    page.fill("input[placeholder='Enter the part number for the legs']", str(order[3]))
    page.fill("#address", str(order[4]))


def submit_order(page):
    """
    Submits the order and handles potential errors.
    """
    while True:
        page.click("#order")
        if page.is_visible(".alert-danger"):
            continue
        if page.is_visible("#receipt"):
            break


def save_receipt_as_pdf(page, order):
    """
    Saves the order receipt as a PDF file, including the robot preview image.
    """
    order_number = order[0]
    receipt_html = page.locator("#receipt").inner_html()

    # Capture the robot preview image
    preview_image_path = f"output/robot_preview_{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=preview_image_path)

    # Add some basic styling to improve PDF appearance
    styled_html = f"""
    <html>
    <head>PrintReceipt</head>
    <body>
        <img src="{preview_image_path}" alt="Robot Preview" style="max-width: 100%; margin-bottom: 20px;">
        {receipt_html}
    </body>
    </html>
    """

    pdf = PDF()
    pdf_path = f"output/receipt_{order_number}.pdf"
    pdf.html_to_pdf(styled_html, pdf_path)

    # Add the robot preview image to the PDF
    pdf.add_files_to_pdf(
        files=[preview_image_path],
        target_document=pdf_path,
        append=True
    )

    # Clean up the temporary image file
    os.remove(preview_image_path)

    print(f"Saved receipt for order {order_number} as PDF with robot preview")

def archive_receipts():
    """
    Archives all receipt PDFs into a ZIP file.
    """
    archive = Archive()
    output_dir = "output"
    zip_file = os.path.join(output_dir, "receipts.zip")
    
    # Create the ZIP file
    archive.archive_folder_with_zip(
        folder=output_dir,
        archive_name=zip_file,
        include="receipt*.pdf"
    )
    
    print(f"Archived all receipts into {zip_file}")

def close_robot_order_website():
    """
    Closes the RobotSpareBin Industries Inc. website.
    """
    print(f"Closing the browser ...")
    pass