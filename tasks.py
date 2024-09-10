from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF

@task
def open_robot_order_website():
    """Insert the sales data for the week and export it as a PDF"""
    navigate_to("https://robotsparebinindustries.com/")
    log_in()
    navigate_to("https://robotsparebinindustries.com/#/robot-order")
    
    fill_form_with_excel_data()
    
    log_out()

def navigate_to(url):
    """Navigates to the given URL"""
    browser.goto(url)

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def log_out():
    """Presses the 'Log out' button"""
    page = browser.page()
    page.click("text=Log out")

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def next_robot():
    page = browser.page()
    page.click("#order-another")

def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return orders

def fill_orders_and_submit_form(order):
    """Fills in the order data and click the 'Submit' button"""
    page = browser.page()
    
    page.select_option("#head", str(order["Head"]))
    page.check(f"#id-body-{order['Body']}")
    page.get_by_placeholder("Enter the part number for the legs").fill(str(order["Legs"]))
    page.fill("#address", order["Address"])
    
    page.click("#order")
    

def screenshot_robot():
    """Take a screenshot of the page"""
    page = browser.page()
    
    while True:
        if page.is_visible("div.alert-danger[role=alert]"):
            page.click("#order")
            continue
        
        locator = page.locator("#robot-preview-image")
        locator.screenshot(path="output/temp_robot.png")
        return

def embed_screenshot_to_receipt(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    order_result_html = page.locator("#receipt").inner_html()
    
    file_list = [
        "output/temp_robot.png:align=center",
        "output/temp_receipt.png",
    ]

    pdf = PDF()
    pdf.html_to_pdf(order_result_html, "output/temp_receipt.png")
    pdf.add_files_to_pdf(files=file_list, target_document=f"output/{order_number}.pdf")

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    orders = get_orders()

    for row in orders:
        close_annoying_modal()
        fill_orders_and_submit_form(row)
        screenshot_robot()
        embed_screenshot_to_receipt(row["Order number"])
        next_robot()
    close_annoying_modal()