from datetime import datetime
from pypdf import PdfWriter,PdfReader
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

def text_seperator(desired_text:str):
    print("-----------------------------------------------------------------------------------------------------------")
    print(desired_text)
    print("-----------------------------------------------------------------------------------------------------------")

def fiscal_year_calc(invoice_total, start_date, end_date):

    date_format = "%m/%d/%Y"
    start_date_obj = datetime.strptime(start_date, date_format)
    end_date_obj = datetime.strptime(end_date, date_format)
    days_amount = (end_date_obj - start_date_obj).days

    fiscal_year_end = "06/30/2025"
    fiscal_year_end_obj = datetime.strptime(fiscal_year_end, date_format)

    time_to_fisc_end = (fiscal_year_end_obj - start_date_obj).days

    current_fy_amount = (invoice_total * time_to_fisc_end)/days_amount
    next_fy_amount = ((days_amount-time_to_fisc_end)*invoice_total)/days_amount
    if current_fy_amount > invoice_total:
        current_fy_amount = invoice_total
        next_fy_amount = 0

    print(f"{round(current_fy_amount, 2)} should be charged to regular expense.")
    print(f"{time_to_fisc_end}")
    print(f"{round(next_fy_amount,2)} should be charged to 1-4-0405-1000.")
    print(f"Just a check:{current_fy_amount + next_fy_amount}")





def surgical_pdf(pdf_path, output_path):


    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        print("Yea it is")
        reader.decrypt('11201')

    gl_coding_page = reader.pages[0]


    invoices = {'Decrypted': [1, 2]}

    for invoice in invoices:
        print(invoice)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        inv_pages = invoices[invoice]

        for inv_page in inv_pages:
            print(inv_page)
            pos = inv_page-1
            writer.add_page(reader.pages[pos])
        new_path = os.path.join(output_path, f"{invoice} I hope.pdf")
        with open(new_path, "wb") as fp:
            writer.write(fp)

class chromeItem:
    def __init__(self,element_type:'str',name:'str'):
        self.element_type = element_type
        self.name = name

    def create_element(self):
        driver = webdriver.Chrome()
        return WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR, f'{self.element_type}[type="text"][name={self.name}]')))

def organize_ckse_txt(save_url,csv_save_path=None):
    vend_ids = []
    vend_names = []
    voucher_ids = []
    voucher_dates = []
    due_dates = []
    ap_types = []
    invoice_nums = []
    yes_checks = []
    voucher_amounts = []
    discount_amounts = []
    take_discs = []
    pay_amounts = []
    gl_amounts = []
    gl_nos = []
    with open(save_url,"r") as ckse_txt:
        ckse_lines = ckse_txt.read().splitlines()
    dividing_line = ckse_lines[7]
    line_list = dividing_line.split(" ")
    # Basically, the max length of text is determined by the max length of those "-----" right under the column names.
    # So I just get the max length, and get all characters that fall within that range.
    vend_id_max = len(line_list[0])
    vend_name_max = len(line_list[1])
    voucher_id_max = len(line_list[2])
    voucher_date_max = len(line_list[3])
    due_date_max = len(line_list[4])
    ap_type_max = len(line_list[5])
    invoice_num_max = len(line_list[6])
    yes_check_max = len(line_list[7])
    voucher_amount_max = len(line_list[8])
    discount_amount_max = len(line_list[9])
    take_disc_max = len(line_list[10])
    pay_amount_max = len(line_list[11])
    gl_amount_max = len(line_list[12])
    gl_no_max = len(line_list[13])
    for ckse_line in ckse_lines:
        try:
            float(ckse_line[:vend_id_max].strip())
        except ValueError:
            # If i find a better way to include the multiple gls, add that condition(s) here
            pass

        # if "Yes" not in ckse_line:
        #     pass
        else:
            pos = 0
            vend_id = ckse_line[:vend_id_max].strip()
            vend_ids.append(vend_id)

            pos+=vend_id_max+1
            vend_name = ckse_line[pos:vend_name_max+pos]
            vend_names.append(vend_name)

            pos+=vend_name_max+1
            voucher_id = ckse_line[pos:voucher_id_max+pos].strip()
            voucher_ids.append(voucher_id)

            pos+=voucher_id_max+1
            voucher_date = ckse_line[pos:voucher_date_max+pos].strip()
            voucher_dates.append(voucher_date)

            pos+=voucher_date_max+1
            due_date = ckse_line[pos:due_date_max+pos].strip()
            due_dates.append(due_date)

            pos+=due_date_max+1
            ap_type = ckse_line[pos:ap_type_max+pos].strip()
            ap_types.append(ap_type)

            pos+=ap_type_max+1
            invoice_num = ckse_line[pos:invoice_num_max+pos].strip()
            invoice_nums.append(invoice_num)

            pos+=invoice_num_max+1
            yes_check = ckse_line[pos:yes_check_max+pos].strip()
            yes_checks.append(yes_check)

            pos+=yes_check_max+1
            voucher_amount = ckse_line[pos:voucher_amount_max+pos].strip().replace(",","")
            if voucher_amount.endswith("-"):
                voucher_amount = f"-{voucher_amount[:-1]}"

            voucher_amount = float(voucher_amount)
            voucher_amounts.append(voucher_amount)

            pos+=voucher_amount_max+1
            discount_amount = ckse_line[pos:discount_amount_max+pos].strip()
            discount_amounts.append(discount_amount)

            pos+=discount_amount_max+1
            take_disc = ckse_line[pos:take_disc_max+pos]
            take_discs.append(take_disc)

            pos+=take_disc_max+1
            pay_amount = ckse_line[pos:pay_amount_max+pos]
            pay_amounts.append(pay_amount)

            pos+=pay_amount_max+1
            gl_amount = ckse_line[pos:gl_amount_max+pos].strip()
            gl_amounts.append(gl_amount)

            pos+=gl_amount_max+1
            gl_no = ckse_line[pos:gl_no_max+pos].strip()
            gl_nos.append(gl_no)

            #print(voucher_id)

    ckse_dict = {'Vendor Id':vend_ids,'Vendor Name':vend_names,'Voucher ID':voucher_ids, 'Voucher Date':voucher_dates, 'Due Date':due_dates,'AP Type':ap_types,'Invoice Number':invoice_nums,
                 'E-Check Status':yes_checks, 'Voucher Amount':voucher_amounts,'Discount Amount':discount_amounts,'Take Disc?':take_discs,'Pay Amount':pay_amounts,'GL Amount':gl_amounts,
                 'GL Number':gl_nos}
    ckse_cleaned_table = pd.DataFrame(ckse_dict)
    ckse_cleaned_table['Vendor Id'] = ckse_cleaned_table['Vendor Id'].astype(str)
    ckse_cleaned_table['Invoice Number'] = ckse_cleaned_table['Invoice Number'].astype(str)
    ckse_cleaned_table['Due Date'] = pd.to_datetime(ckse_cleaned_table['Due Date'], format='%m/%d/%y')
    ckse_cleaned_table['Voucher Date'] = pd.to_datetime(ckse_cleaned_table['Voucher Date'], format='%m/%d/%y')
    ckse_cleaned_table['Invoice Number'] = ckse_cleaned_table['Voucher Amount'].astype(float)
    if csv_save_path is not None:
        ckse_cleaned_table.to_csv(csv_save_path)
    # This returns a clean dataframe of invoices, where all the yes' are assumed to be e-checks and all the blanks are assumed as checks.
    # I can use this on any ckse txt file.
    return ckse_cleaned_table



class Condition:
    def __init__(self, condition, source_df, ckse_type=None, cons_form=None, c_focus=None, t_focus=None, cons_form_op=None):
        self.condition = condition  # The actual condition
        self.source_df = source_df
        self.ckse_type = ckse_type
        self.cons_form = cons_form
        self.column_focus = c_focus
        self.filter_target = t_focus
        self.cons_form_op = cons_form_op
        if str(self.source_df) == 'under_df':
            self.ckse_type = 'Under'
        elif str(self.source_df) == 'over_df':
            self.ckse_type = 'Over'

    def __or__(self, other):

        if other is None:
            # If there's no other condition, return the current condition
            return  self.evaluate()
        if not isinstance(other, Condition):
            return NotImplemented
        # Return a new condition that is the OR of both conditions
        return Condition(None, self.source_df, self.evaluate() | other.evaluate())  # Boolean series combined

    def evaluate(self):
        return eval(self.cons_form,{"self":self,"column_focus":self.column_focus,"filter_target":self.filter_target})

    def confirm_type(self,column_focus:'str',filter_target:'str'):
        self.column_focus = column_focus
        self.filter_target = filter_target
        if self.condition == 'remove':

            self.cons_form = f"self.source_df[column_focus] != filter_target"
            # this lambda function saves the operation and makes it callable
            self.cons_form_op = lambda x, y: x & y
        elif self.condition == 'add':

            self.cons_form = f"self.source_df[column_focus] == filter_target"

            self.cons_form_op = lambda x, y: x | y

def exceptions_splitter(desckse_type,exceptions_tup):
    new_exceptions = None
    for exception in exceptions_tup:
        print(f"Comparing exception type:{exception.ckse_type} to {desckse_type}.")
        if exception.ckse_type == desckse_type:
            if new_exceptions is None:
                new_exceptions = exception.evaluate()
            else:
                new_exceptions = exception.cons_form_op(new_exceptions, exception.evaluate())

        else:
            pass
    if new_exceptions is None:
        print("ALERT")
        new_exceptions = pd.Series(False, index=exceptions_tup[0].source_df.index)
    # print(adj_tup)
    return new_exceptions

def or_and(op,init_condition,condition_obj):
    if op == 'and':
        return init_condition & condition_obj
    elif op == 'or':
        return init_condition | condition_obj

