from loaded_model import invoice_predictor
import os


# Make testing path a directory containing the pdfs you wish to test.
testing_path =r""
folder_path = os.listdir(testing_path)

for file in folder_path:
    if file.endswith(".pdf"):
        full_filename = os.path.join(testing_path, file)
        invoice_predictor(full_filename)


