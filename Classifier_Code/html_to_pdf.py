import pdfkit

import pymupdf
import os

def html_to_pdf(html_string,target_path):
    config = pdfkit.configuration(wkhtmltopdf=r"external_items\wkhtmltopdf\bin\wkhtmltopdf.exe")
    try:
        pdfkit.from_string(html_string,output_path=target_path,configuration=config)
    except OSError:
        print("Likely an image that couldn't be found and converted to html.")
        pass
def merge_pdfs(starting_pdf,email_pdf, output_path, pdf_pass):
    # We are sorting based on the natural numbers, if any in the front.

    pdf_list = [starting_pdf,email_pdf]

    merger = pymupdf.open()



    for pdf in pdf_list:
        with pymupdf.open(pdf) as mfile:

            if mfile.needs_pass:
                mfile.authenticate(pdf_pass)

            mfile.bake(widgets=True)


            merger.insert_pdf(mfile)
            print(mfile)




    merger.save(output_path)


    merger.close()
    #resulting_pdf = pymupdf.open(output_path)

    # for merged_pdf in pdf_list:
    #     resulting_pdf.insert_outline

    print(f"Merged PDF saved to {output_path}")

