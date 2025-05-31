import numpy as np
import os
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.model_selection import cross_val_score
from sklearn.model_selection import StratifiedKFold
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.naive_bayes import GaussianNB
from sklearn.svm import SVC
import joblib
from Classifier_Code.image_processing_funcs import pdf_to_array

# Pdf_dir variable needs to be a directory path that points to subdirectories with unique numerical names that contain the pds or images you wish for it to be trained on.
# In my test, i named the files in the format '(subfolder name - identifier)'. The identifier is just how the files i used have it and how i had training_pull.py name it.

pdf_dir =r""
pdfs_dirs = os.listdir(pdf_dir)

# This is for the Physical core warning you get during KNN model testing.
os.environ['OMP_NUM_THREADS'] = '1'

dataset = []
models = []
results = []
names = []
model_results = {}
size_of_img = 100

for sub_pdf_dir in pdfs_dirs:
    print(sub_pdf_dir)

    joined_pdf_paths = os.path.join(pdf_dir,sub_pdf_dir)
    if os.path.isdir(joined_pdf_paths):
        pdfs = os.listdir(joined_pdf_paths)

        for pdf in pdfs:

            full_path = os.path.join(joined_pdf_paths,pdf)
            the_img_array = pdf_to_array(full_path,size_of_img)

            # plt.imshow(the_img_array, cmap='gray')
            # plt.show()

            dataset.append(np.append(the_img_array, sub_pdf_dir))
            print(full_path)
    else:
        pass

df = pd.DataFrame(dataset)

print(df)
array = df.values
X = array[:,0:size_of_img*size_of_img]
y = array[:,size_of_img*size_of_img]
print(y)

X_train, X_validation, Y_train, Y_validation = train_test_split(X, y, test_size=0.20, random_state=1)

models.append(('LR', LogisticRegression(solver='liblinear', multi_class='ovr')))
models.append(('LDA', LinearDiscriminantAnalysis()))
models.append(('KNN', KNeighborsClassifier()))
models.append(('CART', DecisionTreeClassifier()))
models.append(('NB', GaussianNB()))
models.append(('SVM', SVC(gamma='auto')))

for name, model in models:
    kfold = StratifiedKFold(n_splits=5, random_state=1, shuffle=True)
    # The splits determine how many times you see that annoying warning. With a lot of data, use like 3-4. Try to make sure
    # each label or class has more representations than the splits.
    cv_results = cross_val_score(model, X_train, Y_train, cv=kfold, scoring='accuracy')
    results.append(cv_results)
    names.append(name)
    mean_accuracy = cv_results.mean()
    model_results[name] = mean_accuracy, model
    print('%s: %f (%f)' % (name, mean_accuracy, cv_results.std()))

best_model = max(model_results, key=model_results.get)
print(model_results)
print(best_model)
successful_inv_model = model_results[best_model][1]
print(successful_inv_model)

successful_inv_model.fit(X_train, Y_train)

joblib.dump(successful_inv_model, 'invoice_trained_model.pkl')
print(df)

# This needs to be run everytime you have new data. Logisitic regression has always had the best results for the examples I used, being invoices from different companies
# in pdf format, but other models might be better depending on the data. I made this based on the iris example.