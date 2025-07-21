import os
from openpyxl import load_workbook
import shutil
from PDF_Combinator import merge_pdfs

def directory_search(excel_path, inv_sel_path, worksheet_name="SLED Vouchers", test_state=False):
    # The directory or directories that you want the search to rifle through are put in the path_list variable.
    path_list = [r""]

    searched_dir_list = []

    workbook = load_workbook(excel_path)

    worksheet = workbook[worksheet_name]

    invoice_selection_folder = inv_sel_path

    invoices_all_path = invoice_selection_folder.replace("Invoice Selection","Invoices-All.pdf")

    vou_column = worksheet["C"]
    ven_id_column = worksheet["A"]
    vendor_column = worksheet["B"]

    voucher_list = [v.value for v in vou_column]
    vend_id_list = [vid.value for vid in ven_id_column]
    vend_name_list = [vnm.value for vnm in vendor_column]

    # voucher_list = ["V0356188"]
    # vend_id_list = ["Placeholder"]
    # vend_name_list = ["Dummy"]

    print(voucher_list)
    print("_______________________________________________________________________________________________________________")
    for dir_path in path_list:

        starting_dir_list = os.listdir(dir_path)

        dir_list = []

        for item in starting_dir_list:
            initial_item_path = os.path.join(dir_path, item)
            initial_dir_check = os.path.isdir(initial_item_path)
            if initial_dir_check is True:
                print(initial_item_path)

                dir_list.append(initial_item_path)


        print(len(dir_list))

        while len(dir_list) > 0:

            for found_directory in dir_list:

                for item in os.listdir(found_directory):
                    item_path = os.path.join(found_directory,item)
                    dir_check = os.path.isdir(item_path)
                    if dir_check is True:
                        if item_path != "T:\\":
                            print(item_path)
                            if item_path not in dir_list:
                                dir_list.append(item_path)

                    else:
                        pass
                dir_list.remove(found_directory)
                searched_dir_list.append(found_directory)
                # print(dir_list)
                print(len(dir_list))

        searched_dir_list.append(dir_path)
        # Now I have a list of every folder(or directory) that is within the initial directory.
        # print(searched_dir_list)

        # So for this, up until now, I have been able to retrieve a list of ALL folders and directories contained in each branch
        # of the initial path. Now i can iterate through each one and search for ONLY files. Meaning if its a file, search the name
        # for a match. I need to make sure its not super slow. There's probably quicker ways to iterate through these folders like list comprehension.


    #https://ui5production.brooklaw.edu:5002/ui_webapi/file/downloadFile/CDESSOURCES/032846805222410/CKSE_CDESSOUR_34037?r=6
    #https://ui5production.brooklaw.edu:5002/ui_webapi/file/downloadFile/CDESSOURCES/032846805222410/EFT_T1_010625_112755?r=7

    origin_location = []
    ordered_vou = voucher_list.copy()
    for folder_path in searched_dir_list:
        file_list = os.listdir(folder_path)

        for file in file_list:
            file_path = os.path.join(folder_path, file)
            file_check = os.path.isfile(file_path)
            if file_check is True:
                for vou in voucher_list:
                    pos = voucher_list.index(vou)



                    try:
                        if vou in file:
                            ordered_pos = ordered_vou.index(vou) + 1
                            print(ordered_pos)
                            location_text = f"File for {vou} was found {file_path}"
                            origin_location.append(location_text)
                            voucher_list.remove(vou)
                            print(file_path)
                            target_folder = invoice_selection_folder

                            # new_path = os.path.join(target_folder, file)
                            #
                            # os.rename(new_path)
                            vend_id = vend_id_list[ordered_pos-1]
                            vend_name = vend_name_list[ordered_pos-1].replace("/","")
                            new_file_name = f"{vend_name} ({vend_id}) - {vou}"


                            print(new_file_name)
                            new_path_name = os.path.join(target_folder, f"{ordered_pos}) {vend_name} ({vend_id}) - {vou}.pdf")
                            shutil.copy2(file_path, new_path_name)
                            # os.rename(new_path_name, )
                            print(f"The new path name is {new_path_name}.")
                    except TypeError:
                        if vou is not None:
                                print(vou)

    for new_stuff in voucher_list:
        if not None:
            print(new_stuff)
    print(voucher_list)

    for location in origin_location:
        print(location)

    input("Press enter to combine files in location. Only if satisfied with invoice selection.")
    print("______________________________________________________________________________________________________")
    merge_pdfs(invoice_selection_folder, invoices_all_path)

if __name__ == "__main__":

    print("nice")
