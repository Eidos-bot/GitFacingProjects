import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from openpyxl import load_workbook
import datetime
from selenium.common.exceptions import WebDriverException
import random
import traceback
import dotenv

dotenv.load_dotenv()
# Switch to the normal if the new version is causing trouble
chrome_options = Options()
# chrome_options.add_argument("--headless=new")
attempts = 0
max_attempts = 9000

while attempts < max_attempts:
    try:
        driver = webdriver.Chrome(options=chrome_options)
        attempts += max_attempts+100

    except WebDriverException as e:
        wait_time = random.randrange(1, 20, 1)
        attempts += 1
        print(e)
        print("Damnit! Blocked again!")
        print(f"Waiting for {wait_time} seconds.")
        time.sleep(wait_time)
# FYI
if attempts == max_attempts:
    print("Better luck next time.")
    sys.exit(1)
else:
    input("Press enter to continue")
    pass

start_time = time.time()
today = datetime.date.today().strftime("%m/%d/%Y")
print(today)

# Navigate to one of the colleagues. Use staging for code testing!
production_url = r'https://selfservice.brooklaw.edu/Student/Account/Login'
staging_url = r'https://sstest.brooklaw.edu:8174/Student/Account/Login'

url_to_use = staging_url

username = str(os.getenv("SSVOUCHUSER"))
passw = str(os.getenv("SSVOUCHPASS"))

testing = True

if not testing:
    file_path = input('Raw voucher data excel file path.')
    workbook = load_workbook(file_path)
    worksheet = workbook['Sheet']

    due_date_format = '%m/%d/%Y'

    wait_time = 1
    # This is all the data that needs to be pulled. They can contain lists. WIll likely be an excel file.
    # ----------------------------------------------------------------------------------------------------------------------
    voucher_date = []
    invoice_number = []
    invoice_amount = []
    invoice_date = []
    gl_code = []
    ap_type = []
    # Both the gl_code and ap_type would be the most likely bottlenecks. I can only think of using the GAN for pattern
    # recognition for both. I'll have to test the ocr to be able to read text and locate emails with relevant info.
    vendor_id = []
    desc_text = []
    terms = []

    # This turns the info in each relevant column into a list within the relevant field. There's likely a MUCH better way to
    # do this, but I like the needlessly complex stuff. Makes it easier to follow the basics.
    for i, row in enumerate(worksheet):
        if i == 0:
            continue
        col_a = row[0].value
    # I did the None because when i do ocr i have spaces between invoices. So i just wanted those spaces to be ignored.
        if col_a is not None:
            invoice_number.append(col_a)
        else:
            pass


        col_b = row[1].value
        if col_b is not None:
            invoice_amount.append(col_b)
        else:
            pass

        # I hate datetime objects in Excel. See the result of my hate.
        col_c = row[2].value
        try:
            if col_c is not None:
                col_c_str = str(col_c)
                new_i_date = datetime.datetime.strptime(col_c_str, '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')
                invoice_date.append(new_i_date)
            else:
                pass
        except ValueError:
            if col_c is not None:
                invoice_date.append(col_c)
            else:
                pass

        col_d = row[3].value
        if col_d is not None:
            desc_text.append(col_d)
        else:
            pass

        col_e = row[4].value
        if col_e is not None:
            gl_code.append(col_e)
        else:
            pass

        col_f = row[5].value
        if col_f is not None:
            ap_type.append(col_f)
        else:
            pass

        col_g = row[6].value
        if col_g is not None:
            vendor_id.append(col_g)
        else:
            pass

        col_h = row[7].value
        try:
            if col_h is not None:
                col_h_str = str(col_h)
                new_v_date = datetime.datetime.strptime(col_h_str, '%Y-%m-%d %H:%M:%S').strftime('%m/%d/%Y')
                voucher_date.append(new_v_date)
            else:
                pass
        except ValueError:
            if col_h is not None:
                voucher_date.append(col_h)
            else:
                pass

        col_j = row[9].value
        if col_j is not None:
            terms.append(col_j)
        else:
            pass
# ----------------------------------------------------------------------------------------------------------------------


    print(invoice_amount)
    total_invoices = len(vendor_id)

test_due_date = "07/08/2025"
test_vendor_id = "0092755"
test_invoice_num = "TESTINVOICE"
test_invoice_date = "07/08/2025"

driver.get(url_to_use)
user_input = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][name="UserName"]')))
user_input.send_keys(username)
pass_input = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="password"][name="Password"]')))
pass_input.send_keys(passw)

log_in = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="submit"][id="login-button"]')))
log_in.click()

time.sleep(1)
main_url = url_to_use.replace('Account/Login','ColleagueFinance/Procurement')
driver.get(main_url)


time.sleep(5)
print("Waiting for create tab to be ready.")
create_voucher_tab =  WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.ID, 'tab-create')))
create_voucher_tab.click()

select_entry_type_element =  WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.ID, 'dropdownButtonDocumentType')))
select_entry_type_element.click()
time.sleep(1)
payment_request_selector = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.XPATH, "//a[contains(text(), 'Payment Request')]")))
payment_request_selector.click()

needed_by_date = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.ID, 'date-picker-input-6')))
needed_by_date.send_keys(test_due_date, Keys.ENTER)

vendor_id_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="search"][id="user-search"]')))
vendor_id_element.send_keys(test_vendor_id, Keys.ENTER)
time.sleep(1)
vendor_id_results = driver.find_elements(By.CLASS_NAME, "esg-lookup__list-button")
time.sleep(1)
print(vendor_id_results)
first_result = vendor_id_results[0]
first_result.click()

invoice_input_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="requiredInputInvoiceNo_create"]')))
invoice_input_element.send_keys(test_invoice_num, Keys.ENTER)

invoice_date_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="date-picker-input-7"]')))
invoice_date_element.send_keys(test_invoice_date, Keys.ENTER)

select_ap_type_element =  WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.ID, 'dropdownButtonApType')))
select_ap_type_element.click()
time.sleep(1)
tacp_element = WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'ap-type-menu_1')))
print(tacp_element)
tacp_element.click()

# Need a for loop here for multiple line items.
add_item_element = WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.ID, 'add-item-button')))
add_item_element.click()

time.sleep(500)
