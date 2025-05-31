import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import date
from openpyxl import load_workbook
from datetime import datetime, timedelta
import requests
from bls_functions import organize_ckse_txt, text_seperator
import pandas as pd
import shutil
import os
from PDF_Manipulation_Code.directory_search_base import directory_search

def additional_crit(ckse_amounts):
    additional_criteria_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JSP-OTHER-SEL-FLAG"]')))
    additional_criteria_element.send_keys("Yes")
    additional_criteria_element.send_keys(Keys.ENTER)
    time.sleep(1)
    connective_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JSP-SELECT-CONNECTIVE_1"]')))

    connective_element.send_keys("WITH")
    field_name_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JSP-SELECT-FIELD_1"]')))
    field_name_element.send_keys("VOU.VEN.BAL")
    relation_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JSP-SELECT-REL-OPCODE_1"]')))
    if ckse_amounts == "OVER":
        relation_element.send_keys("GE")
    elif ckse_amounts == "UNDER":
        relation_element.send_keys("LT")

    boundary_value_element = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JSP-SELECT-VALUE_1"]')))
    boundary_value_element.send_keys('"5000"')
    boundary_value_element.send_keys(Keys.ENTER)
    time.sleep(1)

def download_href(href_link, txt_output_file):
    response = requests.get(href_link)

    with open(txt_output_file, 'wb') as txt_file:
        txt_file.write(response.content)


ckseprel_target = r"ckse_test.txt"
chrome_options = Options()
# chrome_options.add_argument("--headless=new")
driver = webdriver.Chrome(chrome_options)
start_time = time.time()
today = date.today().strftime("%m/%d/%Y")

voucher_exclusions = []
vendor_exclusions = []

pren_vendors=[]
pren_vouchers=[]

#If they aren't prenoted yet so won't show up on the ach ckse, put in here. I guess it would also work with vendors/vouchers that normally wouldn't pull.

print(today)
# Navigate to one of the colleagues. Use staging for code testing!
production_url = 'https://ui5production.brooklaw.edu:5002/ui/home/index.html'
staging_url = 'https://ui4test.brooklaw.edu:3443/ui/home/index.html'
username = input('username')
passw = input('password')

date_format = '%m/%d/%y'
"""REMEMBER TO CHANGE"""
days = int(input("Enter 12 for regular ckse and 30 for checks: "))
ckse_template_path = input('CKSE template path:')
ckse_workbook = load_workbook(ckse_template_path)
ckse_sheet = ckse_workbook["COVERSHEET"]

cut_off_date_raw= "05/07/2025"
cut_off_date = cut_off_date_raw.replace("/","")
check_date_raw = input("Enter date in format dd/mm/yyyy: ")
check_date = check_date_raw.replace("/","")
saved_list_name = "Test"
ckse_Type = "B" # P B or E
CKSE_amount = "UNDER"
due_date_obj = datetime.strptime(check_date_raw, date_format) + timedelta(days)
due_date_raw = due_date_obj.strftime("%m/%d/%y")
due_date = due_date_raw.replace("/","")
driver.get(production_url)



# This finds all the window handles, otherwise it would look for the following elements in the original url.
# Then it switches to the second handle which is the new one. Use this anytime multiple windows are open, they should
# be ordered properly in the index.
user_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][name="UserID"]')))
user_input.send_keys(username)
pass_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="password"][name="UserPassword"]')))
pass_input.send_keys(passw)
# ATTENTION:USE '#' FOLLOWED BY THE ID OF THE ELEMENT IF THAT'S ALL YOU HAVE. IT'S PREFERRED TO BE MORE SPECIFIC THOUGH!
time.sleep(1)
# We don;t use webdriver wait here because it would press log in the moment it sees it, but we want it after everything is entered.
# The wait method is fine for situations where you can press it the moment its active, but if an element is active but needs
# something to happen before you activate it, just use the regular find_element.
log_in = driver.find_element("css selector", "#btnLogin")
log_in.click()

form_search_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="form-search"]')))
form_search_element.send_keys("CKSE")
time.sleep(3)
form_search_element.send_keys(Keys.ENTER)
time.sleep(1)
try:
    voucher_selected_element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'button[type="button"][id="popup_lookup_button_0"]')))
    voucher_selected_element.click()
    accept_date_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')))
    accept_date_element.click()
except:
    pass

time.sleep(1)
# Find out how to loop this part in. So far this is just for the initial.
saved_list_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-VAR1"]')))
saved_list_element.send_keys(Keys.CONTROL,'a')
saved_list_element.send_keys(Keys.BACKSPACE)
time.sleep(1)
bank_code_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-VAR17"]')))
bank_code_element.send_keys(Keys.CONTROL, 'a')
bank_code_element.send_keys("T1")
bank_code_element.send_keys(Keys.ENTER)
time.sleep(1)
check_date_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-VAR3"]')))
check_date_element.send_keys(Keys.CONTROL, 'a')
check_date_element.send_keys(check_date)
time.sleep(1)
check_date_element.send_keys(Keys.ENTER)
try:
    accept_date_element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'button[type="button"][id="popup_lookup_button_0"]')))
    accept_date_element.click()
except:
    pass
time.sleep(1)
ckse_type_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-VAR2"]')))
ckse_type_element.send_keys(Keys.CONTROL, 'a')
ckse_type_element.send_keys(ckse_Type)
time.sleep(1)
ckse_type_element.send_keys(Keys.ENTER)
time.sleep(1)

# cut_off_date_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-DATE-VAR2"]')))
# cut_off_date_element.send_keys(Keys.CONTROL, 'a')
# cut_off_date_element.send_keys(cut_off_date)
# print(cut_off_date)
# time.sleep(1)
# due_date_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-DATE-VAR4"]')))
# due_date_element.send_keys(Keys.CONTROL, 'a')
# due_date_element.send_keys(due_date)
# print(due_date)

produce_report_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-VAR15"]')))
produce_report_element.send_keys(Keys.CONTROL, 'a')
produce_report_element.send_keys("Yes")
time.sleep(2)
produce_report_element.send_keys(Keys.ENTER)
time.sleep(1)

# additional_crit(CKSE_amount)

save_button_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'button[type="button"][id="btnFormSave"]')))
save_button_element.click()
print("Save 1")
time.sleep(1)
save_button_element.click()
print("Save 2")
time.sleep(1)
save_button_element.click()
print("Save 3")
time.sleep(1)
try:
    save_button_element.click()
    print("Save 4")
except:
    pass
try:
    time.sleep(1)
    save_button_element.click()
    print("Save 5")
except:
    pass
time.sleep(15)
finish_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'button[type="button"][id="finish-btn"]')))
finish_element.click()
print("Clicked Finish")
time.sleep(1)
save_as_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'button[type="button"][id="reportSaveAs"]')))
save_as_element.click()

download_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID,"fileDownloadBtn")))
download_link = download_button.get_attribute("href")
try:
    download_href(download_link,ckseprel_target)
    print("File Download Complete")
except:
    print("File Download Failed")
time.sleep(2)
try:
    close_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'button[type="button"][id="fileCloseBtn"]')))
    close_element.click()
except:
    print("Failure. Check element id.")
    input("Check element id.")
time.sleep(2)

close_panel_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID,'btn_close_report_browser')))
close_panel_element.click()

time.sleep(2)


def save_ckse(ckse_type, ckse_dataframe,
              template_path=ckse_template_path):
    # template path can be changed.
    path_list = []
    if ckse_type == "ALLCHECKS":
        ckse_path_name = fr"CKSE {ckse_type} {check_date_raw.replace('/', '-')}"
    else:
        ckse_path_name = fr"CKSE {ckse_type} 5K {check_date_raw.replace('/','-')}"
    ckse_excel = fr"Colleague_Code\{ckse_path_name}\{ckse_path_name}.xlsx"
    path_list.append(ckse_excel)
    final_ckse_pdf = ckse_excel.replace(".xlsx",".pdf")
    final_ckse_textfile = ckse_excel.replace(".xlsx",".txt")
    invoice_selection_path = ckse_excel.replace(fr"{ckse_path_name}.xlsx",r"Invoice Selection")
    path_list.append(invoice_selection_path)
    os.makedirs(os.path.dirname(ckse_excel), exist_ok=True)
    os.makedirs(invoice_selection_path, exist_ok=True)
    sled_vouch_df = ckse_dataframe.iloc[:, :3]
    original_ckse_df = ckse_dataframe.iloc[:, 1:]

    shutil.copy2(template_path, ckse_excel)
    sled_name = ckse_path_name.replace(' ', '_')
    single_value_df = pd.DataFrame([[sled_name]])
    with pd.ExcelWriter(ckse_excel, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        ckse_dataframe.to_excel(writer, sheet_name='COVERSHEET', startrow=14, index=False, header=False)
        sled_vouch_df.to_excel(writer, sheet_name='SLED Vouchers', startrow=0, index=False, header=False)
        original_ckse_df.to_excel(writer, sheet_name='Original CKSE', startrow=2, startcol=1, index=False, header=False)
        single_value_df.to_excel(writer, sheet_name='COVERSHEET', startrow=6, startcol=6, index=False, header=False)

    # Being lazy trying to make it all in one
    sled_list = '\n'.join(list(ckse_dataframe['Voucher ID']))
    time.sleep(1)
    form_search_element.send_keys("SLED")
    time.sleep(3)
    form_search_element.send_keys(Keys.ENTER)
    time.sleep(1)

    sled_name_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][name="popup-lookup"]')))

    #Note that colleague will change any spaces into an underscore. So

    sled_name_element.send_keys(sled_name)
    time.sleep(1)
    sled_ok_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')))
    sled_ok_element.click()
    time.sleep(1)
    try:
        sled_ok_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_1"]')))
        sled_ok_element.click()
        time.sleep(1)
    except:
        pass

    sled_detail_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "detail_icon_LIST-VAR1_1")))
    sled_detail_element.click()
    time.sleep(1)
    sled_input_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "multiLine")))
    sled_input_element.send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    sled_input_element.send_keys(sled_list)
    time.sleep(1)
    sled_list_save_button_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "btnCommentDialogSave")))
    sled_list_save_button_element.click()
    time.sleep(1)
    sled_save_button_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "btnFormSave")))
    sled_save_button_element.click()
    time.sleep(1)
    form_search_element.send_keys(Keys.CONTROL, 'a')
    time.sleep(1)
    form_search_element.send_keys("CKSE")
    time.sleep(3)
    form_search_element.send_keys(Keys.ENTER)
    time.sleep(2)
    try:
        voucher_selected_element_final = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')))
        voucher_selected_element_final.click()
        accept_date_element_final = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="popup_lookup_button_0"]')))
        accept_date_element_final.click()
    except:
        pass
    saved_list_element_final = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-VAR1"]')))
    time.sleep(3)
    saved_list_element_final.send_keys(Keys.CONTROL, 'a')
    saved_list_element_final.send_keys(Keys.BACKSPACE)
    saved_list_element_final.send_keys(sled_name)
    time.sleep(1)

    produce_report_element_final = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="text"][id="JS-VAR15"]')))
    produce_report_element_final.send_keys(Keys.CONTROL, 'a')
    produce_report_element_final.send_keys("Yes")
    time.sleep(2)
    produce_report_element_final.send_keys(Keys.ENTER)
    time.sleep(1)

    # additional_crit(CKSE_amount)

    save_button_element_final = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="btnFormSave"]')))
    save_button_element_final.click()
    print("Save 1")
    time.sleep(1)
    save_button_element_final.click()
    print("Save 2")
    time.sleep(1)
    save_button_element_final.click()
    print("Save 3")
    time.sleep(1)
    try:
        save_button_element_final.click()
        print("Save 4")
    except:
        input("Save issue. Please check.")
        pass
    try:
        time.sleep(1)
        save_button_element_final.click()
        print("Save 5")
    except:
        input("Save issue. Please check.")
        pass
    time.sleep(15)
    finish_element_final = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="finish-btn"]')))
    finish_element_final.click()
    print("Clicked Finish")
    time.sleep(1)

    save_as_element_final = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="reportSaveAs"]')))
    save_as_element_final.click()

    download_button_final = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "fileDownloadBtn")))
    download_link_final = download_button_final.get_attribute("href")
    try:
        download_href(download_link_final, final_ckse_textfile)
        print("Text File Download Complete")
    except:
        print("Text File Download Failed")
    time.sleep(2)

    try:
        close_element_final = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="fileCloseBtn"]')))
        close_element_final.click()
    except:
        print("Failure. Check element id.")
        input("Check element id.")
    time.sleep(1)
    export_pdf_element_final = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="reportExport"]')))
    export_pdf_element_final.click()

    create_pdf_final_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="create"]')))
    create_pdf_final_element.click()

    download_button_final = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "fileDownloadBtn")))
    download_link_final = download_button_final.get_attribute("href")
    try:
        download_href(download_link_final, final_ckse_pdf)
        print("PDF File Download Complete")
    except:
        print("PDF File Download Failed")
    time.sleep(2)

    try:
        close_element_final = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'button[type="button"][id="fileCloseBtn"]')))
        close_element_final.click()
    except:
        print("Failure. Check element id.")
        input("Check element id.")
    time.sleep(2)

    close_panel_element_final = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'btn_close_report_browser')))
    close_panel_element_final.click()

    time.sleep(2)
    return path_list




# Remember there's an option to save a csv version of the "raw" form of the ckse. But we dont really need this,
# especially not for the prel.
prel_dataframe = organize_ckse_txt(r"ckse_test.txt")
print(prel_dataframe)
print("Above is the raw prel_dataframe.")

prel_dataframe = prel_dataframe[prel_dataframe['AP Type'] != 'TSRF']
print(prel_dataframe)
print("TSRF has been removed from the prel_dataframe.")
check_df_raw = prel_dataframe[(prel_dataframe['E-Check Status'] != 'Yes')&~((prel_dataframe['Voucher ID'].isin(pren_vouchers))|(prel_dataframe['Vendor Id'].isin(pren_vendors)))]
prel_dataframe = prel_dataframe[(prel_dataframe['E-Check Status'] == 'Yes')|(prel_dataframe['Voucher ID'].isin(pren_vouchers))|(prel_dataframe['Vendor Id'].isin(pren_vendors))]
print("Voucher's without yes to e-check have been removed from the prel_dataframe.")



alt_due_date = due_date_obj.strftime("%m/%d/%Y")

cut_off_alt_raw = datetime.strptime(cut_off_date_raw, "%m/%d/%Y")
alt_cut_off = cut_off_alt_raw.strftime("%m/%d/%Y")

prel_dataframe = prel_dataframe[prel_dataframe['Due Date']<alt_due_date]
print("The due date has been implemented.")

print(prel_dataframe)
vendor_sums = prel_dataframe.groupby('Vendor Id')['Voucher Amount'].sum()
over_vendors = vendor_sums[vendor_sums >= 5000].index
under_vendors = vendor_sums[(0< vendor_sums) & (vendor_sums < 5000)].index

#-----------------------------------------------------------------------------------------------------------------------
# over_df = prel_dataframe[prel_dataframe['Vendor Id'].isin(over_vendors)]
# under_df = prel_dataframe[prel_dataframe['Vendor Id'].isin(under_vendors)]
#
# cond_1 = Condition(condition='add', source_df=under_df)
# cond_2 = Condition(condition='add', source_df=under_df)
# cond_3 = Condition(condition='remove', source_df=over_df)
# cond_4 =  Condition(condition='remove', source_df=over_df)
# cond_5 = Condition(condition='remove', source_df=over_df)
# cond_6 = Condition(condition='remove', source_df=over_df)
#
# cond_1.confirm_type('Vendor Id',"0037651")
# cond_2.confirm_type('Vendor Id',"0444219")
# cond_3.confirm_type('Vendor Id',"0411489")
# cond_4.confirm_type('Vendor Id',"0381778")
# cond_5.confirm_type('Voucher ID',"V0361594")
# cond_6.confirm_type('Voucher ID',"V0361598")
# exceptions = [cond_3,cond_2,cond_5,cond_6]
# under_exceptions = exceptions_splitter("Under",exceptions)
# print("The under exceptions are:")
# print(under_exceptions)
# over_exceptions = exceptions_splitter("Over",exceptions)
# print("The over exceptions are:")
# print(over_exceptions)
# #Changed both to or. Note that and is causing issues. I believe its trying to add the noted vendors or remove them but they aren't seeing them.
# under_df = under_df[or_and('or',under_df['Voucher Date']<=alt_cut_off,under_exceptions)]
# over_df = over_df[or_and('and',over_df['Voucher Date']<=alt_cut_off,over_exceptions)]

#
# under_df = under_df[(under_df['Voucher Date']<=cutoff_date)|(under_df['Vendor Id']=="0037651")|(under_df['Vendor Id']=="0444219")]
# # IT WAS FUCKING AND vvv THIS WAS FUCKING AND!!! THAT'S ^^^ FUCKING BIT OR
# over_df = over_df[(over_df['Voucher Date']<=cutoff_date)&(over_df['Vendor Id']!='0411489')]
#-----------------------------------------------------------------------------------------------------------------------

"""The above was an attempt at using a class I made. It's not necessary for the code, and it honestly makes it far more clunky
than necessary. You notice that one line does the work."""

over_df = prel_dataframe[(prel_dataframe['Vendor Id'].isin(over_vendors) &~ prel_dataframe['Voucher ID'].isin(voucher_exclusions)&~ prel_dataframe['Vendor Id'].isin(vendor_exclusions))]
under_df = prel_dataframe[(prel_dataframe['Vendor Id'].isin(under_vendors) &~ prel_dataframe['Voucher ID'].isin(voucher_exclusions)&~ prel_dataframe['Vendor Id'].isin(vendor_exclusions))]
check_df = check_df_raw[~check_df_raw['Voucher ID'].isin(voucher_exclusions)&~ check_df_raw['Vendor Id'].isin(vendor_exclusions)]

dir_search_targets = []

# text_seperator("Over Dataframe")
# print(over_df)
# print(over_df['Voucher Amount'].sum())
# dir_search_targets.append(save_ckse('OVER',over_df))
# print("Ckse template has been saved.")

text_seperator('Under Dataframe')
print(under_df)
print(under_df['Voucher Amount'].sum())
dir_search_targets.append(save_ckse('UNDER',under_df))
print("Ckse template has been saved.")

# text_seperator("Check Dataframe")
# print(check_df)
# print(check_df['Voucher Amount'].sum())
# dir_search_targets.append(save_ckse('ALLCHECKS',check_df))
# print("Ckse template has been saved.")


# So the SLED lists are saved. Now we go back into colleague, adding some new steps and then doing the ckse process again.
log_out_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'btnLogout_Outer')))
log_out_element.click()

for dir_search_target in dir_search_targets:
    ckse_excel_target = dir_search_target[0]
    inv_election_target = dir_search_target[1]
    directory_search(ckse_excel_target, inv_election_target)
