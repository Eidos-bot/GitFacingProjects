import time
from playwright.sync_api import Playwright, sync_playwright, expect, TimeoutError
import dotenv
import os
from openpyxl import load_workbook
import datetime

dotenv.load_dotenv()

def run(playwright: Playwright, file_path, state=False, colleague_testing=False) -> None:

    workbook = load_workbook(file_path)
    worksheet = workbook['Sheet']

    due_date_format = '%m/%d/%Y'
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
    unique_list = [f"{a}{b}" for a, b in zip(invoice_number, vendor_id)]
    username = str(os.getenv("COLLEAGUEUSER"))
    passw = str(os.getenv("COLLEAGUEPASS"))
    browser = playwright.chromium.launch(headless=state)
    context = browser.new_context()
    page = context.new_page()
    if colleague_testing:
        page.goto("https://sstest.brooklaw.edu:8174/Student/Account/Login")
    else:
        page.goto("https://selfservice.brooklaw.edu/Student/Account/Login")


