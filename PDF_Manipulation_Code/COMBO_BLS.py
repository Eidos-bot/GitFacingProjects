import cv2
import numpy as np
import pytesseract


# Why did i call this combo? Because i solved the issue in parts. First i needed to find out what the data looked at. Embarrassing to type since that's basic
# error handling right? So i found out the text data retrieved from a pytesseract scan is a dictionary of all words, coords and pages where its found.
# This was great! Now the data is not obfuscated. Now its something i recognize. Then i needed to learn how to find the location of a word
# a list. Easy. Then i need to find a specific sequence of that word or words in the list. Got it. Ok. Now i need the coords of that last word and once i get that its simple math or testing.
# I just needed to shift the desired coords using the known words as an anchor. I had to learn all this in parts and finally put it together. Now i have the_finder function as a result.
# The distance parameter is the only one that requires some testing.
# I was going to just get the coords and let the parent code use pytesseract, but i figured since it was looking at the images already, i might as well crop and create the image here. So now
# the end result is a cropped image of the coords created based on the parameters entered. Not perfect but its pretty good. Before i was getting the largest bbox but that was giving me everything.


# images = this is the main source pdf. It should be turned into an image with convert_from_path() from pdf2image. If theres a better one, then use that. But it needs to be an image
# word_search = of course this is the word that you should be looking for. If there are multiple words, it turns it into a list and looks in the text_data list to see if the sequence of this list matches any sequence in there. If not, it uses a dummy roi. It
# uses the last word to get the initial coords. But what if the last word shows multiple times?
# occurrence = spelling this word right was harder than solving this. If you have multiple instances of the word_search, this value specifies which one you want. So if there are two "TOTALS" in text_data, you can choose the 2nd by setting this as 2.
# Don't worry, I know index values start at 0 so that's why I subtract 1 in the beginning :)
# distance = how far is your desired data from the found word? This determines how wide the cropped image is. Just to reiterate, the result of the_finder will be an image. So anything that's typically done with images can be done with a the_finder result.

# To set PATH for Python
pytesseract.pytesseract.tesseract_cmd = r"external_items/Tesseract/tesseract.exe"
def the_finder(images, word_search, occurrence, distance, dummy_roi, dist_from_found_wrd):
    option = occurrence - 1

    image_array = np.array(images)  # Convert PpmImageFile to NumPy array

    gray = cv2.cvtColor(image_array, cv2.COLOR_BGR2GRAY)

    # Perform OCR to extract text and its bounding boxes
    text_data = pytesseract.image_to_data(gray, lang="eng", output_type=pytesseract.Output.DICT)

    sequence = word_search.split()

    indices = []

    for i in range(len(text_data['text']) - len(sequence) + 1):
        if text_data['text'][i:i + len(sequence)] == sequence:
            indices.append(i + len(sequence) - 1)
    print(f"The indices are: {indices}.")
    if not indices:
        new_x, new_y, new_x2, new_y2 = dummy_roi
        print("Can't find word, using dummy roi.")
    else:
        x, y, w, h = text_data['left'][indices[option]], text_data['top'][indices[option]], text_data['width'][indices[option]], text_data['height'][indices[option]]
        # Basically, pytesseract fortunately gives us a dictionary with a text list, a leftmost list and topmost.
        # Those are the only coordinates in the dictionary. To get the rightmost and bottommost coords I just add the width and height.
        # So that's what x2 and y2 are. Now we will designate a large "safe area" to in a certain direction.
        x2 = w + x
        y2 = h + y
        new_x, new_y, new_x2, new_y2 = x2 + 10 + dist_from_found_wrd, y, x2 + distance, y2+5
    # this creates the new image based on the new coordinates. The margin is small but you can change it if needed
    margin = 5
    cropped_image = image_array[max(0, new_y - margin):min(image_array.shape[0], new_y2 + margin), max(0, new_x - margin):min(image_array.shape[1], new_x2 + margin)]

    return cropped_image

# Future plans, not really high priority will be to make it multi-directional. Right now, the only coords found are
# to the right. In the future I'd like to add top, bottom, and left.


# This is used if you have a roi image that has multiple levels of text, and you want it read from left to right.
# It will organize the text properly because otherwise it won't be consistent.
def ocr_text_realignment(roi_image):
    text_data_1 = pytesseract.image_to_data(roi_image, lang="eng", output_type=pytesseract.Output.DICT)
    top_list = text_data_1['top']
    text_list = text_data_1['text']
    left_list = text_data_1['left']
    new_list = list(zip(top_list, left_list, text_list))
    # We are sorting the new list of those three elements by the top value, which makes sure things on the same x plane
    # will be next to each other. And by
    sorted_list = sorted(new_list, key=lambda x: (x[0], x[1]))
    new_text_list = []
    for tup in sorted_list:
        # This is taking the sorted text and putting it in the new list for it to then be converted to a string. I KNOW,
        # IT'S UGLY!
        text = tup[2]
        new_text_list.append(text)
    new_text_str_raw = ' '.join(new_text_list)
    new_text_str = ' '.join(new_text_str_raw.split())
    return new_text_str
