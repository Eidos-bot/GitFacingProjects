import numpy as np
from pdf2image import convert_from_path

def resize_image(image, target_size):
    if isinstance(target_size, int):
        target_size = (target_size, target_size)

    return image.resize(target_size)

def pdf_to_array(pdf_path, size_of_image):


    converted_img = convert_from_path(pdf_path,
                                      poppler_path=r"external_items/poppler-24.08.0/Library/bin")

    con_img = converted_img[0]
    resized_img = resize_image(con_img, size_of_image)
    gray = resized_img.convert("L")
    img_array = np.array(gray)


    # plt.imshow(the_img_array, cmap='gray')
    # plt.show()


    return img_array
