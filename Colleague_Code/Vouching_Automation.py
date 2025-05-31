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
production_url = r'https://ui5production.brooklaw.edu:5002/ui/home/index.html'
staging_url = r'https://ui4test.brooklaw.edu:3443/ui/index.html'

username = input('username')
passw = input('password')

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

driver.get(production_url)

print(invoice_amount)
total_invoices = len(vendor_id)
# link_element = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#launchlink")))
# link_element.click()
# # This clicks the link to open the colleague window.
#
# time.sleep(4)
# all_windows = driver.window_handles
# driver.switch_to.window(all_windows[1])
# This finds all the window handles, otherwise it would look for the following elements in the original url.
# Then it switches to the second handle which is the new one. Use this anytime multiple windows are open, they should
# be ordered properly in the index.
user_input = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][name="UserID"]')))
user_input.send_keys(username)
pass_input = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="password"][name="UserPassword"]')))
pass_input.send_keys(passw)
# ATTENTION:USE '#' FOLLOWED BY THE ID OF THE ELEMENT IF THAT'S ALL YOU HAVE. IT'S PREFERRED TO BE MORE SPECIFIC THOUGH!
time.sleep(1)
log_in = driver.find_element("css selector", "#btnLogin")
log_in.click()
time.sleep(10)
# It's this long so it can finish authenticating. Can make it shorter, though you risk slow internet messing things up.
form_search_element = WebDriverWait(driver, 200).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="form-search"]')))
form_search_element.send_keys("VOUM")
time.sleep(3)
voucher_dict = {}
try:
    for value in invoice_number:
        position = invoice_number.index(value)
        # This is where the invoice loop will start.
        form_search_element.send_keys(Keys.ENTER)

        voucher_lookup_add_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="popup-lookup"]')))
        voucher_lookup_add_element.send_keys("A")
        voucher_lookup_add_element.send_keys(Keys.ENTER)
        time.sleep(1)
        # For the below, is asked again regardless of the date. This is just to get rid of the date prompt so the loop settles.
        # if position > 0:
        #
        #     accept_voucherdate_element = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR,
        #                                                      'button[type="button"][id="popup_lookup_button_0"]')))
        #     time.sleep(1)
        #     accept_voucherdate_element.click()
        #     print("Successful click.")
        #     time.sleep(1)
        # else:
        #     pass
        # Explanation: The only time this is needed is when you need to go into the voucher date to change it. The only
        # time that happens is if you are using a different date than the current date. So this if/else checks for that.
        # if voucher_date[position] != today:
        #     if position > 0:
        #         pass
        #     else:
        #         print("AT THE PROBLEM")
        #
        #         accept_voucherdate_element = WebDriverWait(driver, 100).until(EC.visibility_of_element_located(("css selector",
        #                                                          'button[type="button"][id="popup_lookup_button_0"]')))
        #         accept_voucherdate_element.click()
        #         time.sleep(2)
        # else:
        #     pass

        voucher_date_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-DATE"]')))
        voucher_date_element.send_keys(voucher_date[position])
        print(f"The voucher date is: {voucher_date[position]}.")
        voucher_date_element.send_keys(Keys.ENTER)
        time.sleep(1)
        # THIS IS NECESSARY IF THE FY ISNT CLOSED
        # if position > 0 and voucher_date[position] == today:
        #     pass
        # else:
        #     accept_voucherdate_element = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')))
        #     accept_voucherdate_element.click()


        # FOR SOME REASON MY PUNY BRAIN CANT UNDERSTAND I NEED TO FIND IT AGAIN FOR IT TO BE RECOGNIZED
        # If info is entered before any prompt is answered, even if it shows up in the element, the system won't recognize it.


        invoice_number_element = WebDriverWait(driver, 200).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-DEFAULT-INVOICE-NO"]')))
        time.sleep(1)
        actual_inv = invoice_number[position]
        invoice_number_element.send_keys(actual_inv)
        time.sleep(1)
        invoice_number_element.send_keys(Keys.ENTER)
        print(f"The invoice number is: {actual_inv}.")
        # Be sure to name anything with a data source that will be the same as an element
        time.sleep(2)
        invoice_date_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-DEFAULT-INVOICE-DATE"]')))
        invoice_date_element.send_keys(invoice_date[position])
        print(f"The invoice date is: {invoice_date[position]}.")
        # Since there's no prompt for date's, let's try to not do a sleep time until voucher done element.
        invoice_total_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-INVOICE-AMT"]')))
        invoice_total_element.send_keys(invoice_amount[position])
        print(f"The total invoice amount is: {invoice_amount[position]}.")
        time.sleep(3)
        vendor_id_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-VENDOR"]')))
        vendor_id_element.send_keys(vendor_id[position])
        print(f"The vendor id is: {vendor_id[position]}.")
        ap_type_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-AP-TYPE"]')))
        # The ap type field remains populated with the type of the last invoice. So control all will make sure its removed.
        ap_type_element.send_keys(Keys.CONTROL, 'a')
        time.sleep(1)
        print(position, len(ap_type))

        ap_type_element.send_keys(ap_type[position])
        ap_type_element.send_keys(Keys.TAB)
        print(f"The ap type is: {ap_type[position]}.")
        time.sleep(5)
        try:

            accept_ap_type_element = driver.find_element("css selector", 'button[type="button"][id="popup_lookup_button_0"]')
            accept_ap_type_element.click()
            voucher_done_element = WebDriverWait(driver, 100).until(
                ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VAR3"]')))
            voucher_done_element.send_keys(Keys.CONTROL, 'a')
            time.sleep(2)
        except:
            print("Accepted AP Type.")
            pass
        time.sleep(1)
        due_date_obj = datetime.datetime.strptime(invoice_date[position], due_date_format)
        adj_due_date = due_date_obj + datetime.timedelta(days=terms[position])
        due_date = adj_due_date.strftime('%m/%d/%Y')
        due_date_element = WebDriverWait(driver, 100).until(
                ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-DUE-DATE"]')))
        due_date_element.send_keys(Keys.CONTROL, 'a')
        time.sleep(1)
        due_date_element.send_keys(due_date)
        print(f"Due date is {due_date}.")
        time.sleep(1)
        voucher_done_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VAR3"]')))
        voucher_done_element.send_keys(Keys.CONTROL, 'a')
        time.sleep(2)
        done_or_no = "Yes"
        voucher_done_element.send_keys(done_or_no)
        time.sleep(1)
        voucher_done_element.send_keys(Keys.ENTER)
        time.sleep(1)

        # if done_or_no == "Yes":
        #     accept_voucherdate_element = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')))
        #     accept_voucherdate_element.click()
        # else:
        #     pay_voucher_element = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="VOU-PAY-FLAG"]')))
        #     pay_voucher_element.send_keys(Keys.CONTROL, 'a')
        #     time.sleep(1)
        #     pay_voucher_element.send_keys(done_or_no)
        #     time.sleep(1)
        #     pay_voucher_element.send_keys(Keys.ENTER)
        #     pass
        line_items_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, "#detail_icon_VAR5")))
        line_items_element.click()
        time.sleep(1)
        voil_item_list_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, "#detail_icon_VOU-ITEMS-ID_1")))
        voil_item_list_element.click()
        time.sleep(1)

        # Loop for gl would start here if there are multiple codes
        desc_detail_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, "#detail_icon_ITM-DESC_1")))
        desc_detail_element.click()


        # You might be wondering why there are so many time.sleep(1)s, or if you are an actual coder, you think it's ugly.
        # The plan is to switch to checking to see if the elements are there and once they are, to interact with them. It will
        # reduce the wait times and everything will be smoother. The sleeps are just for troubleshooting right now, so I can see
        # if everything that is supposed to be clicked, is. Once I'm confident enough to make it headless, ill switch methods.

        desc_text_area_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, "#multiLine")))
        desc_text_area_element.send_keys(desc_text[position])
        print(f"The description is: {desc_text[position]}.")

        save_desc_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="btnCommentDialogSave"]')))
        save_desc_element.click()

        gl_area_amount_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="ITM-VOU-PRICE"]')))
        # This usually equals the invoice_amount value but only if just one gl code is used. Will make a modifier for that.
        # Like if the list of gl codes is >1, do the requisite calculations. That should also affect the amount of times the gl
        # loop is done.
        time.sleep(1)
        gl_area_amount_element.send_keys(invoice_amount[position])
        gl_area_amount_element.send_keys(Keys.ENTER)
        # Will need another list for gl codes
        time.sleep(1)
        print(f"GL Code {gl_code[position]} will have {invoice_amount[position]} assigned to it.")
        gl_area_quantity_element = driver.find_element("css selector", 'input[type="text"][id="ITM-VOU-QTY"]')
        gl_area_quantity_element.send_keys(1)
        # Usually one so no need for a variable.
        gl_area_quantity_element.send_keys(Keys.ENTER)
        time.sleep(1)
        gl_area_code_element = driver.find_element("css selector", 'input[type="text"][id="ITM-VOU-GL-NO_1"]')
        gl_area_code_element.send_keys(gl_code[position])
        gl_area_code_element.send_keys(Keys.ENTER)
        print(f"The GL code is {gl_code[position]}.")

        time.sleep(1)
        try:
            # Project Code
            time.sleep(2)
            project_popup_element = driver.find_element("css selector", "#person-search-grid")
            print("Select Project!!!")
            input("Enter when ready.")

        except:
            print("No Project")
            pass
        time.sleep(2)
        gl_area_percent_element = driver.find_element("css selector", 'input[type="text"][id="ITM-VOU-GL-PCT_1"]')
        gl_area_percent_element.send_keys(100)
        # Usually 100%.
        time.sleep(1)
        gl_area_amount2_element = driver.find_element("css selector", 'input[type="text"][id="ITM-VOU-GL-AMT_1"]')
        time.sleep(1)
        gl_area_amount2_element.send_keys(Keys.CONTROL, 'a', invoice_amount[position])

        try:
            time.sleep(2)
            # This is for the overbudget ok button. I made it all one line
            # because it wasn't clicking right for some reason.
            # It works like this though.
            driver.find_element(By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]').click()
            time.sleep(1)
            try:
                initial_ok = driver.find_element(By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')
                initial_ok.click()
                time.sleep(2)
                driver.find_element(By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_1"]')
                print("Its found. ")
            except:
                pass
            time.sleep(2)
            print("Overbudget")
            input("Press ok and reset to press yes.")

            over_budget_yes_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_1"]')))
            print("It's over budget and yes element found.")
            time.sleep(1)
            # Please solve this issue!!!!!!!!
            """Lets try to make a separate definition that's only active here for over-budget. Actually a lambda."""
            # over_budget_yes_element.click()
            #
            # time.sleep(15)
            over_budget_auth_element = driver.find_element("css selector", 'input[type="text"][id="popup-lookup"]')
            over_budget_auth_element.send_keys("CDessources")
            time.sleep(2)
            over_budget_auth_ok_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR,
                                                              'button[type="button"][id="popup_lookup_button_0"]')))

            over_budget_auth_ok_element.click()
            time.sleep(1)
            over_budget_auth_ok_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR,
                                                              'button[type="button"][id="popup_lookup_button_0"]')))
            over_budget_auth_ok_element.click()
        except:
            print("Not Overbudget")
            pass
        # Need this pause so we can wait for the amount to "settle". Otherwise it will save before it registers you accepted the overbudget.
        time.sleep(1)
        save_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="btnFormSave"]')))
        save_element.click()
        time.sleep(1)
        # This is where the GL loop should end
        gl_area_cancel_element = WebDriverWait(driver, 100).until(ec.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="btnFormCancel"]')))
        gl_area_cancel_element.click()
        print("Exiting GL section.")
        time.sleep(1)
        save_element.click()
        time.sleep(1)
        voucher_done_element = driver.find_element("css selector", 'input[type="text"][id="VAR3"]')
        voucher_done_element.send_keys(Keys.ENTER)
        # if done_or_no == "Yes":
        #     accept_voucherdate_element = WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')))
        #     accept_voucherdate_element.click()
        #     time.sleep(1)
        # else:
        #     pass
        voucher_number_element = driver.find_element("css selector", "#DISPLAY-VOUCHERS-ID_headerlabel").text
        print(f"The voucher number for {vendor_id[position]} invoice {invoice_number[position]} is: {voucher_number_element}.")
        print(f"The position is {position+1} of {total_invoices}.")
        print("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------")

        voucher_dict[voucher_number_element]=actual_inv
        input("Review and press enter if satisfactory.")
        save_element.click()
        time.sleep(1)
        # This is where the invoice loop ends.
except Exception as e:

    print(f"Something went wrong:{e}.")
    traceback.print_exc()
else:
    print("Successful completion.")
finally:
    os.startfile(file_path)

print(f"Done with invoices. The vouchers are {voucher_dict}")
end_time = time.time()
elapsed_time = end_time - start_time
true_elapsed_time = str(datetime.timedelta(seconds=elapsed_time))
print("Elapsed time: ", true_elapsed_time)
time.sleep(30)
driver.quit()
