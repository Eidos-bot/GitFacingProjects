import pytesseract
from COMBO_BLS import the_finder

# To set PATH for Python
pytesseract.pytesseract.tesseract_cmd = r"external_items/Tesseract/tesseract.exe"
def clean_text(text):
    return ''.join(filter(str.isprintable, text))
# Function to clean non-printable characters from text
def Slimming_ROI(roi, specialroi, targettext, image, occurrence, distance, dummy_roi, dist_from_found_wrd):
    if roi == (specialroi):
        # Use the alternate ROI obtained from a separate definition
        roi_image = the_finder(image, targettext, occurrence, distance, dummy_roi, dist_from_found_wrd)
        # roi_image = image.crop(alternate_roi)  # the imported function

        # Perform OCR on the ROI image and clean the extracted text
        ocr_text = pytesseract.image_to_string(roi_image)
        raw_text = clean_text(ocr_text)
        replaced_text = raw_text.replace(")", ' ').replace("}", ' ').replace("]", ' ').replace("(", '-').replace("|",' ')
        print(f"Changed {raw_text} to {replaced_text}")
        return replaced_text
