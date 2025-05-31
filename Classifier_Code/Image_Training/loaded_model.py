import joblib
import numpy as np
import PIL.PpmImagePlugin
from Classifier_Code.image_processing_funcs import pdf_to_array, resize_image



# Now we know how to add new data.

# Confidence threshold is of course how high I want it to be before it gives predictions. This has been a decent number as some might look too similar since companies can tend to use similar
# invoicing formats. If that's the case, you need to add more examples of the thing you want to classify.
def invoice_predictor(classification_target,pre_encoded:bool=False,show_low_confidence:bool=False,confidence_threshold:float=99.9):
    loaded_model = joblib.load('invoice_trained_model.pkl')
    prex_new_data = None

    # Set pre_encoded to True if you're encoding it before entering.
    if pre_encoded is True:
        if len(classification_target) > 100:
            raise Exception("Flatten invoice image array to 100 columns. AKA the encoded image is larger than 100 by 100 pixels.")
        elif isinstance(classification_target,np.ndarray) is False:
            raise Exception("This is not a numpy array, set pre_encoded to False or remove it.")

        else:
            prex_new_data = classification_target

    elif pre_encoded is False:
        if isinstance(classification_target, PIL.PpmImagePlugin.PpmImageFile):
            # If this is a PIL image object (expected from a list of them), it will adjust accordingly.
            # This is if you want to iterate over pages in a single pdf. I would have deprecated, but what if you get a pdf
            # that has a bunch of different invoices in one? This helps identify each page individually.
            # Alternatively, if you were confident that every page is relevant, I could adjust this for that.

            resized_img = resize_image(classification_target, 100)
            gray = resized_img.convert("L")
            prex_new_data = np.array(gray)
        else:
            prex_new_data = pdf_to_array(classification_target, 100).flatten()
    # print(prex_new_data)
    # Note that the original training data was already "flattened" by being in a dataframe. So we need to flatten
    # new data ourselves since we aren't doing that. The change after is just a column numbering the row. In our case, 1.
    # Remember that in training data and here, a row is just one image.

    x_new_data = prex_new_data.reshape(1, -1)



    new_prediction = loaded_model.predict(x_new_data)
    probabilities = loaded_model.predict_proba(x_new_data)

    fixed_probs = np.round(probabilities*100,4).astype(str).tolist()

    # print(probabilities)

    predicted_class = str(new_prediction[0])

    highest_prob = max(fixed_probs[0])
    if float(highest_prob)>confidence_threshold or show_low_confidence is True:
        # The show low confidence is if you want to see what low confident rating the thing has.
        if pre_encoded is True:

            print(f"Probability the image is {predicted_class} is {highest_prob}%.")
        elif pre_encoded is False:
            print(f"Probability {classification_target} is {predicted_class} is {highest_prob}%.")
        return predicted_class
    else:
        print(f"Probability of {classification_target} is too low for confident classification.")
        return None


