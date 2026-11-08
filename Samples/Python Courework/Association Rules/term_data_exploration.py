# -*- coding: utf-8 -*-
"""
Created on Sun Feb 11 13:07:44 2024

@author: ems31
"""

import pandas as pd

# Load the first file
project_data = pd.read_csv('D:/Gradschool/SWENG545/Term Project/Projectdata.csv')

# Load the second file
course_list = pd.read_csv('D:/Gradschool/SWENG545/Term Project/Course_List.csv')

# Identify the column names, the number of rows and columns, and unique values for both files

# Project Data
project_data_info = {
    "column_names": project_data.columns.tolist(),
    "number_of_rows": project_data.shape[0],
    "number_of_columns": project_data.shape[1],
    "unique_values": {column: project_data[column].nunique() for column in project_data.columns}
}

# Course List
course_list_info = {
    "column_names": course_list.columns.tolist(),
    "number_of_rows": course_list.shape[0],
    "number_of_columns": course_list.shape[1],
    "unique_values": {column: course_list[column].nunique() for column in course_list.columns}
}

project_data_info, course_list_info


# Identify any missing values in the Project Data file

# Check for missing values and count them
missing_values_count = project_data.isnull().sum()

# Filter rows with at least one missing value
rows_with_missing_values = project_data[project_data.isnull().any(axis=1)]

# Count the number of records with missing values
num_records_with_missing = rows_with_missing_values.shape[0]

missing_values_count, num_records_with_missing

# Output the records that are missing the 'coursename' values
rows_with_missing_coursename = rows_with_missing_values[['Student name', 'semester new', 'coursename']]
rows_with_missing_coursename


# Find, count, and output any duplicate records in the Project Data file

# Find duplicate records (considering all columns)
duplicates = project_data[project_data.duplicated()]

# Count duplicate records
num_duplicates = duplicates.shape[0]

num_duplicates, duplicates




from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import linear_kernel

# Function to calculate Jaccard similarity
def jaccard_similarity(query, document):
    intersection = set(query).intersection(set(document))
    union = set(query).union(set(document))
    return len(intersection) / len(union)

# Preprocess course names: lower case and remove extra spaces
course_names = course_list['Course_Name '].str.lower().str.strip()

# Calculate similarity matrix
vectorizer = TfidfVectorizer().fit_transform(course_names)
cosine_similarities = linear_kernel(vectorizer, vectorizer)

# Find pairs with high similarity but not identical
similar_pairs = []
for i in range(cosine_similarities.shape[0]):
    for j in range(i+1, cosine_similarities.shape[0]):
        if 0 < cosine_similarities[i, j] < 1:  # Exclude identical pairs and non-similar
            # Calculate Jaccard similarity for further filtering
            jaccard_sim = jaccard_similarity(course_names.iloc[i].split(), course_names.iloc[j].split())
            if jaccard_sim > 0.7:  # Threshold for "very, very close"
                similar_pairs.append((course_names.iloc[i], course_names.iloc[j], jaccard_sim))

# Convert to DataFrame for easier viewing
similar_pairs_df = pd.DataFrame(similar_pairs, columns=['Course Name 1', 'Course Name 2', 'Jaccard Similarity'])

similar_pairs_df



project_data_cleaned = project_data.drop_duplicates().dropna(subset=['coursename'])

# Deduplicate 'coursename' before similarity check
unique_coursenames = project_data_cleaned['coursename'].str.lower().str.strip().drop_duplicates()

# Calculate similarity matrix for unique 'coursename'
vectorizer_unique_coursenames = TfidfVectorizer().fit_transform(unique_coursenames)
cosine_similarities_unique_coursenames = linear_kernel(vectorizer_unique_coursenames, vectorizer_unique_coursenames)

# Find pairs with high similarity but not identical, for unique 'coursename' only
similar_pairs_unique_coursenames = []
for i in range(cosine_similarities_unique_coursenames.shape[0]):
    for j in range(i+1, cosine_similarities_unique_coursenames.shape[0]):
        if 0 < cosine_similarities_unique_coursenames[i, j] < 1:  # Exclude identical pairs and non-similar
            # Calculate Jaccard similarity for further filtering
            jaccard_sim = jaccard_similarity(unique_coursenames.iloc[i].split(), unique_coursenames.iloc[j].split())
            if jaccard_sim > 0.7:  # Threshold for "very, very close"
                similar_pairs_unique_coursenames.append((unique_coursenames.iloc[i], unique_coursenames.iloc[j], jaccard_sim))

# Convert to DataFrame for easier viewing
similar_pairs_unique_coursenames_df = pd.DataFrame(similar_pairs_unique_coursenames, columns=['Course Name 1', 'Course Name 2', 'Jaccard Similarity'])

similar_pairs_unique_coursenames_df




# Deduplicate 'Student name' before similarity check
unique_student_names = project_data_cleaned['Student name'].str.lower().str.strip().drop_duplicates()

# Calculate similarity matrix for unique 'Student name'
vectorizer_unique_student_names = TfidfVectorizer().fit_transform(unique_student_names)
cosine_similarities_unique_student_names = linear_kernel(vectorizer_unique_student_names, vectorizer_unique_student_names)

# Find pairs with high similarity but not identical, for unique 'Student name' only
similar_pairs_unique_student_names = []
for i in range(cosine_similarities_unique_student_names.shape[0]):
    for j in range(i+1, cosine_similarities_unique_student_names.shape[0]):
        if 0 < cosine_similarities_unique_student_names[i, j] < 1:  # Exclude identical pairs and non-similar
            # Calculate Jaccard similarity for further filtering
            jaccard_sim = jaccard_similarity(unique_student_names.iloc[i].split(), unique_student_names.iloc[j].split())
            if jaccard_sim > 0.2:  # Threshold for "very, very close"
                similar_pairs_unique_student_names.append((unique_student_names.iloc[i], unique_student_names.iloc[j], jaccard_sim))

# Convert to DataFrame for easier viewing
similar_pairs_unique_student_names_df = pd.DataFrame(similar_pairs_unique_student_names, columns=['Student Name 1', 'Student Name 2', 'Jaccard Similarity'])

similar_pairs_unique_student_names_df



# First, standardize the 'coursename' and 'Course_Name' columns for matching
project_data_cleaned['coursename_cleaned'] = project_data_cleaned['coursename'].str.lower().str.strip()
course_list['Course_Name_Cleaned'] = course_list['Course_Name '].str.lower().str.strip()

# Perform an outer join on the cleaned 'coursename' and 'Course_Name' to find unmatched courses
merged_data = pd.merge(project_data_cleaned, course_list, left_on='coursename_cleaned', right_on='Course_Name_Cleaned', how='outer', indicator=True)

# Identify courses from Projectdata not found in Course_List (and vice versa)
unmatched_courses = merged_data.loc[merged_data['_merge'] != 'both']

# Count unique unmatched courses from Projectdata
unique_unmatched_courses_projectdata = unmatched_courses.loc[unmatched_courses['_merge'] == 'left_only', 'coursename_cleaned'].unique()
num_unique_unmatched_courses_projectdata = len(unique_unmatched_courses_projectdata)

# Output some of the unmatched courses from Projectdata
sample_unmatched_courses_projectdata = unique_unmatched_courses_projectdata[:10]  # Show first 10 as a sample

num_unique_unmatched_courses_projectdata, sample_unmatched_courses_projectdata



# Re-import pandas and reload the data due to code execution state reset
import pandas as pd


# Standardize the 'Course_Name' column for identifying duplicates
course_list['Course_Name_Cleaned'] = course_list['Course_Name '].str.lower().str.strip()

# Identify the duplicate course name in the Course List
duplicate_course_name = course_list[course_list['Course_Name_Cleaned'].duplicated(keep=False)]['Course_Name '].unique()

duplicate_course_name



# Re-import necessary libraries and reload data due to execution state reset
import pandas as pd
import numpy as np
from sklearn.metrics.pairwise import linear_kernel
from sklearn.feature_extraction.text import TfidfVectorizer



# Standardize the 'Course_Name' column for similarity check
course_list['Course_Name_Cleaned'] = course_list['Course_Name '].str.lower().str.strip()

# Deduplicate course names
unique_course_names = course_list['Course_Name_Cleaned'].unique()

# Vectorize the unique course names using TF-IDF
vectorizer = TfidfVectorizer()
tfidf_matrix = vectorizer.fit_transform(unique_course_names)

# Calculate cosine similarities
cosine_similarities = linear_kernel(tfidf_matrix, tfidf_matrix)

# Function to convert cosine similarity to Jaccard similarity
def cosine_to_jaccard(cosine_similarity):
    # Jaccard Similarity = 1 - cosine distance
    return 1 - (1 - cosine_similarity) / 2

# Calculate Jaccard similarities from cosine similarities
jaccard_similarities = cosine_to_jaccard(cosine_similarities)

# Find pairs with high Jaccard similarity but not identical
similar_pairs = []
for i in range(len(unique_course_names)):
    for j in range(i + 1, len(unique_course_names)):
        if 0 < jaccard_similarities[i, j] <= 1:  # Adjusted to include all non-identical pairs
            similar_pairs.append((unique_course_names[i], unique_course_names[j], jaccard_similarities[i, j]))

# Filter for high similarity scores, but not identical
high_similarity_pairs = [pair for pair in similar_pairs if pair[2] > 0.7 and pair[2] < 1]

# Convert to DataFrame for easier viewing, showing top 10 for brevity
similar_pairs_df = pd.DataFrame(high_similarity_pairs, columns=['Course Name 1', 'Course Name 2', 'Jaccard Similarity']).head(10)

similar_pairs_df

