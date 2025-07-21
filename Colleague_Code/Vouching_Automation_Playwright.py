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
        page.goto("https://ui4test.brooklaw.edu:3443/ui/index.html")
    else:
        page.goto("https://ui5production.brooklaw.edu:5002/ui/index.html")

    with page.expect_popup() as page1_info:
        page.get_by_role("link", name="Click here to launch UI").click()
    page1 = page1_info.value
    page1.get_by_role("textbox", name="User name", exact=True).click()
    page1.get_by_role("textbox", name="User name", exact=True).fill(username)
    page1.get_by_role("textbox", name="User Password").click()
    page1.get_by_role("textbox", name="User Password").fill(passw)
    page1.get_by_role("button", name="Log In").click()
    page1.get_by_role("textbox", name="Search for a form").click()
    page1.get_by_role("textbox", name="Search for a form").fill("VOUM")
    page1.get_by_role("textbox", name="Search for a form").press("Enter")
    for invoice_item in unique_list:
        position = unique_list.index(invoice_item)
        time.sleep(1)
        page1.get_by_role("textbox", name="Lookup prompt for Voucher").fill("a")
        page1.get_by_role("textbox", name="Lookup prompt for Voucher").press("Enter")
        page1.get_by_role("button", name="Y", exact=True).click()
        time.sleep(1)
        page1.get_by_role("textbox", name="Maintainable field Voucher Date is a Date field and is").fill(voucher_date[position])
        page1.get_by_role("textbox", name="Maintainable field Voucher Date is a Date field and is").press("Enter")
        try:
            page1.get_by_role("button", name="Y", exact=True).click(timeout=5)
        except TimeoutError:
            print("No confirmation needed.")
            pass
        page1.get_by_role("textbox", name="Maintainable field Vendor ID").click()
        page1.get_by_role("textbox", name="Maintainable field Vendor ID").fill(vendor_id[position])
        page1.get_by_role("textbox", name="Maintainable field Vendor ID").press("Enter")
        time.sleep(1)
        try:
            not_set_ap_type_warning_locator = page1.locator( "[id='popup_lookup_button_0']")
            not_set_ap_type_warning_locator.click(timeout=100)
        except TimeoutError:
            print("No not set ap type warning.")
            pass

        invoice_num_locator = page1.locator( "[id='VOU-DEFAULT-INVOICE-NO']")
        invoice_num_locator.fill(str(invoice_number[position]))
        time.sleep(1)
        invoice_num_locator.press("Tab")
        time.sleep(1)
        try:
            duplicate_invoice_warning_locator = page1.locator( "[id='popup_lookup_button_0']")
            duplicate_invoice_warning_locator.click(timeout=100)
        except TimeoutError:
            print("No duplicate invoice warning.")
            pass


        page1.get_by_role("textbox", name="Maintainable field Invoice Date is a Date field").click()
        page1.get_by_role("textbox", name="Maintainable field Invoice Date is a Date field").fill(invoice_date[position])


        invoice_amt_locator = page1.locator( "[id='VOU-INVOICE-AMT']")
        invoice_amt_locator.wait_for(state="visible")
        total_invoice_amt = str(invoice_amount[position])
        invoice_amt_locator.fill(total_invoice_amt)
        expect(invoice_amt_locator).to_have_value(total_invoice_amt)
        invoice_amt_locator.press("Tab")
        time.sleep(1)

        ap_type_locator = page1.locator("[id='VOU-AP-TYPE']")
        ap_type_locator.fill(ap_type[position])
        ap_type_locator.press("Tab")

        time.sleep(1)
        try:
            not_set_ap_type_warning_locator = page1.locator( "[id='popup_lookup_button_0']")
            not_set_ap_type_warning_locator.click(timeout=100)
        except TimeoutError:
            print("No not set ap type warning.")
            pass

        value = invoice_amt_locator.input_value().strip()
        if value != total_invoice_amt:

            print("VOU-INVOICE-AMT value error", value)
            time.sleep(1)
        print("You are here.")
        time.sleep(1)
        page1.locator("#detail_icon_VAR5").get_by_role("img").click()
        page1.locator("#detail_icon_VOU-ITEMS-ID_1").get_by_role("img").click()
        page1.locator("#detail_icon_ITM-DESC_1").get_by_role("img").click()
        page1.get_by_role("textbox", name="Multiline editor content").click()
        page1.get_by_role("textbox", name="Multiline editor content").fill(desc_text[position])
        time.sleep(1)
        page1.get_by_label("Save").click()
        page1.get_by_role("textbox", name="Maintainable field Price is a").click()
        page1.locator( "[id='ITM-VOU-PRICE']").fill(total_invoice_amt)
        page1.get_by_role("textbox", name="Maintainable field Price is a").press("Tab")
        time.sleep(1)
        page1.get_by_role("textbox", name="Maintainable field Quantity is a Calculator field", exact=True).press("Tab")
        time.sleep(1)

        gl_code_locator = page1.locator( "[id='ITM-VOU-GL-NO_1']")
        gl_code_locator.fill(str(gl_code[position]))
        gl_code_locator.press("Tab")
        proj_id = "250-HEATX-25"
        proj_id.replace("250","")
        time.sleep(1)
        if gl_code_locator.input_value().strip().startswith("4-0"):
            proj_id_locator = page1.get_by_role("cell", name=f"{proj_id}")
            time.sleep(1)
            try:
                proj_id_locator.click(timeout=50)
                time.sleep(1)
                page1.locator("#btnNonFormSearchResultOpenBottom").click()

            except TimeoutError:
                print("Going to next page")
                page1.get_by_role("button", name="Next Page (PageDown)").click()
                proj_id_locator.click(timeout=500)
                page1.locator("#btnNonFormSearchResultOpenBottom").click()
            except TimeoutError:
                print("Going to next page")
                page1.get_by_role("button", name="Next Page (PageDown)").click()
                proj_id_locator.click(timeout=500)
                page1.locator("#btnNonFormSearchResultOpenBottom").click()
            except TimeoutError:
                print("Going to next page")
                page1.get_by_role("button", name="Next Page (PageDown)").click()
                proj_id_locator.click(timeout=500)
                page1.locator("#btnNonFormSearchResultOpenBottom").click()
            except TimeoutError:
                print("Proj ID is not available/valid.")
        # I would include a bool to the voucher dictionary to decide this path
        else:
            time.sleep(1)
            try:
                time.sleep(1)
                print("Exiting project id selection")
                exit_projid_selection_locator = page1.locator( "[id='non-form-close']")
                exit_projid_selection_locator.wait_for(state="visible")
                exit_projid_selection_locator.click(timeout=500)
                time.sleep(2)
                no_proj_id_confirmation_locator = page1.locator( "[id='popup_lookup_button_0']")
                time.sleep(1)
                no_proj_id_confirmation_locator.click(timeout=500)
            except TimeoutError:
                print("Problem with exiting")

        page1.get_by_role("textbox", name="Window ITM.VOU.GL.NO row 1 Maintainable field Foreign Amount is a Calculator").click()
        time.sleep(1)
        page1.get_by_role("button", name="Save", exact=True).click()
        time.sleep(1)
        page1.get_by_role("button", name="Cancel", exact=True).click()
        time.sleep(1)
        page1.get_by_role("button", name="Save", exact=True).click()
        time.sleep(1)
        page1.get_by_role("textbox", name="Maintainable field Voucher Done is").fill("N")
        page1.get_by_role("textbox", name="Maintainable field Voucher Done is").press("Tab")
        readiness_check = page1.get_by_role("textbox", name="Maintainable field Voucher Done is").input_value().strip()
        print(readiness_check)
        if readiness_check != "No":
            print("Marked as ready.")
            page1.get_by_role("button", name="Y", exact=True).click()

        voucher_label_locator = page1.locator( "[id='DISPLAY-VOUCHERS-ID_headerlabel']")
        voucher_id = voucher_label_locator.inner_text()
        print("voucher_id", voucher_id)
        time.sleep(1)
        page1.get_by_role("button", name="Save", exact=True).click()
        print("Successful save!")

    page1.locator("#popup_lookup_button_1").click()
    page1.get_by_role("link", name="Log Out").click()
    page1.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playw:
    run(playw,r"C:\Users\christopher.dessourc\BLS OCR Target\Voucher Data Raw.xlsx")
