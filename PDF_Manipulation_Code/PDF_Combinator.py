import os
import pymupdf
from pypdf import PdfWriter,PdfReader
import re





def natural_sort_key(s, _nsre=re.compile('([0-9]+)')):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(_nsre, s)]
""" This is much much slower and less capable than the merge pdfs function. Don't use it!!! """
def combine_pdfs(pdf_folder, output_path):
    # Create a PdfMerger object
    merger = PdfWriter()


    # Get a list of all PDF files in the specified folder
    files = [f for f in os.listdir(pdf_folder) if os.path.isfile(os.path.join(pdf_folder, f))]
    pdf_files = sorted(files, key=natural_sort_key)
    # Sort the PDF files to ensure they are merged in the desired order
    print(pdf_files)

    # Append each PDF file to the merger object
    for pdf_file in pdf_files:

        pdf_path = os.path.join(pdf_folder, pdf_file)
        print(pdf_path + pdf_file)
        reader = PdfReader(pdf_path)
        if reader.is_encrypted:
            reader.decrypt("")
            for page in reader.pages:
                merger.add_page(page)

        else:
            merger.append(pdf_path)



    # Write the combined PDF to the output path
    merger.write(output_path)
    merger.close()

    print(f"Combined PDF saved to {output_path}")

# file_to_add = r""
# # word_to_pdf(r"")
# main_path = r""
# directory_list = os.listdir(main_path)
# x = 1
# total = len(directory_list)
# for files in directory_list:
#
#     if files.endswith(".pdf") and files is not os.path.basename(file_to_add):
#         pdf_path = main_path + "/" + files
#
#
#         new_file = []
#
#         new_file.append(pdf_path)
#         new_file.append(file_to_add)
#         merger = PdfWriter()
#
#         for pdf in new_file:
#             merger.append(pdf)
#         pdf_path_new = pdf_path.replace(".pdf", " -.pdf")
#         merger.write(pdf_path_new)
#         merger.close()
#
#         print(f"{x} of {total}. {pdf_path_new}")
#         x += 1

# Example usage
# combine_pdfs(r"",
#           r"")

# combine_pdfs(r"",r"")

def merge_pdfs(pdf_folder, output_path):
    # We are sorting based on the natural numbers, if any in the front.
    pdf_list = sorted(os.listdir(pdf_folder), key=natural_sort_key)
    #  The toc
    toc = []
    merger = pymupdf.open()
    page_num = 1
    for pdf in pdf_list:
        fixed_pdf = os.path.join(pdf_folder, pdf)

        with pymupdf.open(fixed_pdf) as mfile:
            num_pages = len(mfile)
            pdf_name = pdf.replace(".pdf","")
            mfile.bake(widgets=True)

            merger.insert_pdf(mfile,show_progress=1)
            print(mfile)
            toc.append((1,pdf_name,page_num))
            for i in range(num_pages):
                # We are making the start of the page the first level (Main Bookmark) but also including it in the sub hierarchies.
                # The second level is all we need after the main. However, if we eventually find a way to know what pages are
                # Approvals vs invoices, we can make another hierarchy specific to that. For now, we just increment the
                # pages by i, the number of pages per invoice.
                toc.append((2,pdf_name,page_num+i))


            page_num += num_pages

            print(f"The number of pages in this file is {num_pages}")
    # doesn't save info in form fields
    # Ignore the yellow it works. Use the below to see if true.
    # print("set_toc" in dir(merger))
    print(toc)
    merger.set_toc(toc)
    merger.save(output_path)



    #resulting_pdf = pymupdf.open(output_path)

    # for merged_pdf in pdf_list:
    #     resulting_pdf.insert_outline

    print(f"Merged PDF saved to {output_path}")



