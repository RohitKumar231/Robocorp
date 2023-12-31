from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF


browser = Selenium()
@task
def robot_spare_bin_python():
    browser.open_available_browser("https://robotsparebinindustries.com/")
    login()
    download_excel_file()
    # fill_and_submit_sales_form()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()


def login():
    browser.input_text('id:username','maria')
    browser.input_text('id:password','thoushallnotpass')
    browser.click_button('Log in')

def fill_and_submit_sales_form(sales_rep):
    browser.input_text_when_element_is_visible('id:firstname', sales_rep["First Name"])
    browser.input_text('id:lastname',sales_rep["Last Name"])
    browser.select_from_list_by_value("id:salestarget", str(sales_rep["Sales Target"]))
    browser.input_text('id:salesresult',str(sales_rep["Sales Target"]))
    browser.click_button('Submit')  
    
def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)
    

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_and_submit_sales_form(row)

def collect_results():
    """Take a screenshot of the page"""
    browser.capture_page_screenshot(filename="output/sales_summary.png")


def export_as_pdf():
    """Export the data to a pdf file"""
    sales_results_html = browser.get_element_attribute('id:sales-results','outerHTML')

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")