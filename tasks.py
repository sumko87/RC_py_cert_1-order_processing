from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
import time
from RPA.Excel.Files import Files
from RPA.PDF import PDF


@task
def order_processing():
    log_into_robot_spare_bin()
    download_input_file()
    orders = read_xlsx_file()
    for order in orders:
        capture_order(order)
    collect_results()
    export_as_pdf()
    log_out
    



def log_into_robot_spare_bin():
    #open browser and navigate to robot_spare_bin, and log into the app
    browser.configure(browser_engine="chrome",slowmo=100)
    browser.goto("https://robotsparebinindustries.com/")
    time.sleep(5)
    page=browser.page()
    page.fill("#username","maria")
    page.fill("#password","thoushallnotpass")
    page.click("button:text('Log in')")
    time.sleep(5)

def download_input_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx",target_file="input.xlsx", overwrite=True)   

def read_xlsx_file():
    files=Files()
    files.open_workbook("input.xlsx")
    orders = files.read_worksheet_as_table(header=True)
    files.close_workbook
    return orders


def capture_order(order):
    print(order["Last Name"])
    page=browser.page()
    page.fill("#firstname", order["First Name"])
    page.fill("#lastname", order["Last Name"])
    page.select_option("#salestarget", str(order["Sales Target"]))
    page.fill("#salesresult", str(order["Sales"]))
    page.click("button:text('Submit')")

def collect_results():
    page=browser.page()
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf():
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")    


def log_out():
    page=browser.page()
    page.click("text=Log out")