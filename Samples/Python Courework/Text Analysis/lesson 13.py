# -*- coding: utf-8 -*-
"""
Created on Tue Jun 12 16:37:25 2018

@author: Leo
"""


 	

# Import the necessary packages.
import nltk
import pandas as pd
import os
import numpy as np


# Load data:

path = 'E:/GoogleDrive/PSU/DAAN862/Course contents/Lesson 13/SentenceCorpus//labeled_articles'
# Change working directory
os.chdir(path)
file_names = os.listdir()  # Create a list of all file names
# file_names.pop(0)  # since I am using google drive, there is an additional file in the folder.
file_names[:5]

from sklearn.feature_extraction.text import CountVectorizer
count_vect = CountVectorizer(stop_words = 'english')
porter = nltk.PorterStemmer()

# Take a look at the text
file = open(file_names[0], 'r')
text1 = file.read()
file.close()
text1[:100]

df = pd.DataFrame()
for i in range(1, len(file_names)):
    # Open the file and read text from it
    file = open(file_names[i], 'r')
    text = file.read()
    file.close()
    # Remove useless sentences
    text = text = text.replace('### abstract ###\n','')
    text = text.replace('### introduction ###\n', '')
    text = text.replace('CITATION', '')    
    text = text.replace('SYMBOL', '')           
    # get all sentences in the text.
    sentences = nltk.sent_tokenize(text) 
    for j in range(len(sentences)):
        # get all words in each sentece
        words = nltk.word_tokenize(sentences[j])
        if words[0] in ['MISC', 'OWNX','CONT', 'AIMX', 'BASE']:
            # The first word in the category of the sentence.
            type_of_sent = words[0]
            # Normalize the rest words
            words = [porter.stem(w) for w in words[1:]]
            # count the freq of the words in the text
            data = count_vect.fit_transform(words)
            # uniqe word list
            unique_words = count_vect.get_feature_names()
            # Word frequency is stored in a condense matrix
            freq = np.sum(data.toarray(), axis = 0)
            # Reshape the frequency from (n, ) to (1, n)
            freq = np.array(freq).reshape(1, (len(freq)))
            # Create a temporary data frame to store the word list and frequency
            temp = pd.DataFrame(freq, columns =unique_words)
            # Assign the category to each sentence
            temp['CLASS'] = type_of_sent
            # merge new file data to the final dataframe
            df = pd.concat([df, temp], ignore_index=True, sort = True)
    print(i)

# Handle missing values
df.isnull().sum().sum()
df = df.fillna(0)
df.isnull().sum().sum()
# Check dataframe information.
df.shape
df.CLASS.value_counts()

# Naive bayes
from sklearn.model_selection import train_test_split 
from sklearn import metrics
from sklearn.naive_bayes import GaussianNB
colnames = list(df)
colnames.remove('CLASS')
X = df[colnames]
y = df.CLASS

X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.33, random_state=134)

NB = GaussianNB()
NB.fit(X_train, y_train)

NB_pred_train = NB.predict(X_train)
NB_pred_test = NB.predict(X_test)
metrics.accuracy_score(y_train, NB_pred_train )
metrics.accuracy_score(y_test, NB_pred_test )
metrics.confusion_matrix(y_train, NB_pred_train )
metrics.confusion_matrix(y_test, NB_pred_test )


# Support Vector Machine


os.chdir('E:/GoogleDrive/PSU/DAAN862/Course contents/Lesson 13')
dbword = pd.read_csv('dbworld_bodies.csv')
dbword.shape
dbword.CLASS.value_counts()
dbword.isnull().sum().sum()

n = dbword.shape[1]
X = dbword.iloc[:, 0:(n-2)]
y = dbword.CLASS

X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.33, random_state=34)

from sklearn import svm
svm_model = svm.SVC(kernel = 'linear')
svm_model.fit(X_train, y_train)

svm_pred_train = svm_model.predict(X_train)
svm_pred_test = svm_model.predict(X_test)
metrics.accuracy_score(y_train, svm_pred_train)
metrics.accuracy_score(y_test, svm_pred_test)
metrics.confusion_matrix(y_train, svm_pred_train).ravel()
metrics.confusion_matrix(y_test, svm_pred_test).ravel()

