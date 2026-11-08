# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 06:01:59 2024

@author: ems31
"""

import pandas as pd

# Load the CSV file to examine its contents
file_path_apriori = 'D:/\GradSchool/SWENG545/Term Project/Final_Apriori.csv'
data_apriori = pd.read_csv(file_path_apriori)

# Split and expand the Antecedent and Consequent columns, then stack them for a uniform series of electives
antecedent_series = data_apriori['Antecedent'].str.split(', ').apply(pd.Series).stack().reset_index(drop=True)
consequent_series = data_apriori['Consequent'].str.split(', ').apply(pd.Series).stack().reset_index(drop=True)

# Use pd.concat instead of append for compatibility with newer pandas versions
electives_series = pd.concat([antecedent_series, consequent_series])

# Calculate the support for each elective by counting its occurrences as both Antecedent and Consequent
elective_support_counts = electives_series.value_counts()

# Convert to a DataFrame for better readability
elective_support_df = elective_support_counts.reset_index()
elective_support_df.columns = ['Elective', 'SupportCount']

# Sort electives by their support count in descending order
elective_support_df_sorted = elective_support_df.sort_values(by='SupportCount', ascending=False)

# Display the top 20 electives by support count
top_electives = elective_support_df_sorted.head(20)
print(top_electives)

############### CAN SAY support of 25?

# Load the electives file to examine its contents
file_path_electives = 'D:/\GradSchool/SWENG545/Term Project/Final_Electives_Preproc.csv'
data_electives = pd.read_csv(file_path_electives)

# Display the first few rows of the dataset to understand its structure
data_electives.head()


# Calculate the frequency of each elective course
elective_frequency_counts = data_electives['Course_Name'].value_counts().reset_index()
elective_frequency_counts.columns = ['Course_Name', 'Frequency']

# Sort the electives by their frequency in descending order
elective_frequency_counts_sorted = elective_frequency_counts.sort_values(by='Frequency', ascending=False)

elective_frequency_counts_sorted.head(20)



###################
from mlxtend.preprocessing import TransactionEncoder
from mlxtend.frequent_patterns import association_rules
from mlxtend.frequent_patterns import apriori

transactions = data_electives.groupby('Student name')['Course_Name'].apply(list).tolist()


encoder = TransactionEncoder()
encoded_matrix = encoder.fit(transactions).transform(transactions)
df_encoded = pd.DataFrame(encoded_matrix, columns=encoder.columns_)


frequent_itemsets = apriori(df_encoded, min_support=0.1, use_colnames=True)

rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=0.1)

print(rules[['antecedents', 'consequents', 'support', 'confidence', 'lift']])


##############

def calculate_support(itemset, transactions):
    """Calculate the support of an itemset within transactions."""
    count = sum(1 for transaction in transactions if itemset.issubset(transaction))
    return count / len(transactions)

def generate_candidates(last_frequent, length):
    return list(set([item for sublist in last_frequent for item in sublist]))


def apriori_with_supports(transactions, min_support):
    """Apriori algorithm that also returns support values for each itemset."""
    # Convert transactions to a list of sets for efficient processing
    transactions = list(map(set, transactions))
    
    # Initial single itemsets
    single_items = set(item for transaction in transactions for item in transaction)
    current_l = [frozenset([item]) for item in single_items]
    
    # Store frequent itemsets with their supports
    itemsets_with_supports = []
    
    k = 2
    while current_l:
        # Calculate support for current itemsets
        itemset_supports = [(itemset, calculate_support(itemset, transactions)) for itemset in current_l]
        # Filter itemsets by min_support and add to the list of frequent itemsets
        current_l = [itemset for itemset, support in itemset_supports if support >= min_support]
        itemsets_with_supports.extend(itemset_supports)
        
        # Generate new candidates: create combinations of the current frequent itemsets
        current_candidates = set([item for subset in current_l for item in subset])
        current_l = [frozenset(items) for items in itertools.combinations(current_candidates, k)]
        
        k += 1
    
    # Filter the final list by min_support one more time to ensure all itemsets meet the criteria
    final_itemsets_with_supports = [(itemset, support) for itemset, support in itemsets_with_supports if support >= min_support]
    
    return final_itemsets_with_supports


import itertools

# Assuming 'transactions' has been defined as your list of course lists per student
frequent_itemsets_with_support = apriori_with_supports(transactions, min_support=0.01)

# Display the frequent itemsets and their supports
for itemset, support in frequent_itemsets_with_support:
    print(f"Itemset: {set(itemset)}, Support: {support}")


# Run the modified Apriori function to get itemsets with their support values
frequent_itemsets_with_support = apriori_with_supports(transactions, min_support=0.01)

# Filter for itemsets that are single courses
single_course_supports = [(itemset, support) for itemset, support in frequent_itemsets_with_support if len(itemset) == 1]

# Sort single courses by their support in descending order
single_course_supports_sorted = sorted(single_course_supports, key=lambda x: x[1], reverse=True)

# Display the sorted single courses and their supports
for itemset, support in single_course_supports_sorted:
    print(f"Course: {list(itemset)[0]}, Support: {support:.4f}")



def calculate_support_count(itemset, transactions):
    """Calculate the absolute support count of an itemset within transactions."""
    return sum(1 for transaction in transactions if itemset.issubset(transaction))

def apriori_with_support_counts(transactions, min_support_count):
    """Apriori algorithm that returns itemsets with their absolute support counts."""
    transactions = list(map(set, transactions))
    single_items = set(item for transaction in transactions for item in transaction)
    current_l = [frozenset([item]) for item in single_items]
    
    # Store itemsets with their absolute support counts
    itemsets_with_counts = []
    
    k = 2
    while current_l:
        itemset_counts = [(itemset, calculate_support_count(itemset, transactions)) for itemset in current_l]
        # Filter itemsets by min_support_count
        current_l = [itemset for itemset, count in itemset_counts if count >= min_support_count]
        itemsets_with_counts.extend(itemset_counts)
        
        # Generate new candidates
        current_candidates = set([item for subset in current_l for item in subset])
        current_l = [frozenset(items) for items in itertools.combinations(current_candidates, k)]
        k += 1
    
    # Filter the final list by min_support_count to ensure all itemsets meet the criteria
    final_itemsets_with_counts = [(itemset, count) for itemset, count in itemsets_with_counts if count >= min_support_count]
    
    return final_itemsets_with_counts


# Adjust min_support_count based on your dataset and needs
min_support_count = 2  # Example minimum support count

# Calculate itemsets with their absolute support counts
itemsets_with_support_counts = apriori_with_support_counts(transactions, min_support_count=min_support_count)

# Display itemsets and their support counts
for itemset, count in itemsets_with_support_counts:
    print(f"Itemset: {set(itemset)}, Support Count: {count}")
    
# Ensure itemsets_with_support_counts is calculated as before
# itemsets_with_support_counts = apriori_with_support_counts(transactions, min_support_count=min_support_count)

# Sort itemsets by their support count in descending order
itemsets_with_support_counts_sorted = sorted(itemsets_with_support_counts, key=lambda x: x[1], reverse=True)

# Display the top 20 itemsets
print("Top 20 itemsets by support count:")
for itemset, count in itemsets_with_support_counts_sorted[:20]:
    # Convert each itemset from a frozenset to a more readable list
    courses = list(itemset)
    print(f"Itemset: {courses}, Support Count: {count}")






# Correcting the approach to filter and identify the top 25% based on highest support and confidence

# Re-apply Step 1 correctly: Filter for Top 25% by Support
support_cutoff = data_apriori['RelativeItemSetSupport%'].quantile(0.50)
filtered_by_support = data_apriori[data_apriori['RelativeItemSetSupport%'] >= support_cutoff]

# Apply Step 2 correctly: Within the filtered group, select the Top 25% by Confidence
confidence_cutoff = filtered_by_support['RuleConfidence%'].quantile(0.50)
final_selection = filtered_by_support[filtered_by_support['RuleConfidence%'] >= confidence_cutoff]

# Display the size of the final selection and the first few rows to understand the selected rules
final_selection_count = len(final_selection)
final_selection.head(), final_selection_count












