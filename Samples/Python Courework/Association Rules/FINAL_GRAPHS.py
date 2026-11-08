# -*- coding: utf-8 -*-
"""
Created on Sat Apr 13 14:29:12 2024

@author: ems31
"""

import pandas as pd

# Load the Apriori results file
file_path = 'D:/\GradSchool/SWENG545/Term Project/Final_Apriori_After_lift_and_freq.csv'
apriori_data = pd.read_csv(file_path)

# Display the first few rows of the dataframe
apriori_data.head()

import networkx as nx
import matplotlib.pyplot as plt

# Extract unique consequents
unique_consequents = apriori_data['Consequent'].unique()

# Create a directed graph
G = nx.DiGraph()

# Add nodes
for consequent in unique_consequents:
    G.add_node(consequent)

# Add edges
for i in unique_consequents:
    for j in unique_consequents:
        if i != j:
            # Calculate average lift between i and j
            subset = apriori_data[(apriori_data['Consequent'] == i) & (apriori_data['Antecedent_List'].str.contains(j))]
            if not subset.empty:
                avg_lift = subset['RuleLift'].mean()
                G.add_edge(i, j, weight=avg_lift)

# Draw the graph
plt.figure(figsize=(12, 12))
pos = nx.circular_layout(G)  # positions for all nodes in a circle
edges = G.edges(data=True)


# Correcting the edge weight extraction
weights = [attr['weight'] for u, v, attr in edges]

# Redraw the graph with correct weights handling
plt.figure(figsize=(12, 12))
nx.draw_networkx(G, pos, with_labels=True, node_size=3000, node_color='skyblue', font_size=9, 
                 edge_color=weights, width=4, edge_cmap=plt.cm.Blues, arrowstyle='-|>', arrowsize=20)

# Show plot
plt.title('Network Graph of Consequents with Edge Weights Based on Average Rule Lift')
plt.show()







############################### Full Apriori ItemSet network graph

import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors

# Assuming the file has been loaded and path is known
apriori_data = pd.read_csv('D:/\GradSchool/SWENG545/Term Project/Final_Apriori.csv')

# Extract unique consequents
unique_consequents = apriori_data['Consequent'].unique()

# Create a directed graph
G_all_supports = nx.DiGraph()

# Add nodes
for consequent in unique_consequents:
    G_all_supports.add_node(consequent)

# Add all individual edges with item set support as weights
for _, row in apriori_data.iterrows():
    antecedents = row['Antecedent'].split(', ')
    consequent = row['Consequent']
    support = row['ItemSetSupport']
    for antecedent in antecedents:
        if antecedent in unique_consequents:
            G_all_supports.add_edge(antecedent, consequent, weight=support)

# Define the color map and the normalization for item set support values using 'Greys'
cmap_greys = plt.cm.Greys
norm_greys = mcolors.Normalize(vmin=min([attr['weight'] for _, _, attr in G_all_supports.edges(data=True)]),
                               vmax=max([attr['weight'] for _, _, attr in G_all_supports.edges(data=True)]))

# Draw the graph
plt.figure(figsize=(14, 14))
pos_all_supports = nx.circular_layout(G_all_supports)
nx.draw_networkx(G_all_supports, pos_all_supports, with_labels=True, node_size=3500, node_color='lightblue', font_size=10,
                 edge_color=[cmap_greys(norm_greys(weight)) for _, _, weight in G_all_supports.edges(data='weight')],
                 width=4, arrowstyle='-|>', arrowsize=20)

# Create a colorbar with the normalization and cmap_greys
sm_greys = plt.cm.ScalarMappable(cmap=cmap_greys, norm=norm_greys)
sm_greys.set_array([])
cbar_greys = plt.colorbar(sm_greys)
cbar_greys.set_label('Itemset Support')

# Show plot with color bar
plt.title('Network Graph of Course Associations with Itemset Support Values')
plt.show()



#################################


################################# Network graph of refined Apriori

import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
# Load the Apriori results file
file_path = 'D:/GradSchool/SWENG545/Term Project/Final_Apriori_After_lift_and_freq.csv'
apriori_data = pd.read_csv(file_path)

unique_consequents = apriori_data['Consequent'].unique()

# Create a directed graph
G_all_supports = nx.DiGraph()

# Add nodes
for consequent in unique_consequents:
    G_all_supports.add_node(consequent)

# Add all individual edges with item set support as weights
for _, row in apriori_data.iterrows():
    antecedents = row['Antecedent_List'].split(', ')
    consequent = row['Consequent']
    support = row['ItemSetSupport']  # Using ItemSetSupport instead of RuleLift
    for antecedent in antecedents:
        if antecedent in unique_consequents:
            G_all_supports.add_edge(antecedent, consequent, weight=support)

# Define the color map and the normalization for item set support values using 'Greys'
cmap_greys = plt.cm.Greys
norm_greys = mcolors.Normalize(vmin=min([attr['weight'] for _, _, attr in G_all_supports.edges(data=True)]),
                               vmax=max([attr['weight'] for _, _, attr in G_all_supports.edges(data=True)]))

# Draw the graph
plt.figure(figsize=(14, 14))
pos_all_supports = nx.circular_layout(G_all_supports)
nx.draw_networkx(G_all_supports, pos_all_supports, with_labels=True, node_size=3500, node_color='lightblue', font_size=10,
                 edge_color=[cmap_greys(norm_greys(weight)) for _, _, weight in G_all_supports.edges(data='weight')],
                 width=4, arrowstyle='-|>', arrowsize=20)

# Create a colorbar with the normalization and cmap_greys
sm_greys = plt.cm.ScalarMappable(cmap=cmap_greys, norm=norm_greys)
sm_greys.set_array([])
cbar_greys = plt.colorbar(sm_greys)
cbar_greys.set_label('Itemset Support')

# Show plot with color bar
plt.title('Network Graph of Consequents with Itemset Support')
plt.show()












import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors


# Extract unique consequents
unique_consequents = apriori_data['Consequent'].unique()

# Create a directed graph
G_all_lifts = nx.DiGraph()

# Add nodes
for consequent in unique_consequents:
    G_all_lifts.add_node(consequent)

# Add all individual edges with lifts as weights
for _, row in apriori_data.iterrows():
    antecedents = row['Antecedent_List'].split(', ')
    consequent = row['Consequent']
    lift = row['RuleLift']
    for antecedent in antecedents:
        if antecedent in unique_consequents:
            G_all_lifts.add_edge(antecedent, consequent, weight=lift)

# Define the color map and the normalization for lift values using 'Greys'
cmap_greys = plt.cm.Greys
norm_greys = mcolors.Normalize(vmin=min([attr['weight'] for _, _, attr in G_all_lifts.edges(data=True)]), vmax=max([attr['weight'] for _, _, attr in G_all_lifts.edges(data=True)]))

# Draw the graph
plt.figure(figsize=(14, 14))
pos_all_lifts = nx.circular_layout(G_all_lifts)
nx.draw_networkx(G_all_lifts, pos_all_lifts, with_labels=True, node_size=3500, node_color='lightblue', font_size=10,
                 edge_color=[cmap_greys(norm_greys(weight)) for _, _, weight in G_all_lifts.edges(data='weight')],
                 width=4, arrowstyle='-|>', arrowsize=20)

# Create a colorbar with the normalization and cmap_greys
sm_greys = plt.cm.ScalarMappable(cmap=cmap_greys, norm=norm_greys)
sm_greys.set_array([])
cbar_greys = plt.colorbar(sm_greys)
cbar_greys.set_label('Lift Value')

# Show plot with color bar
plt.title('Network Graph of Consequents with Individual Lift Values for Each Rule')
plt.show()





# Step 1: Confirm the unique consequents
unique_consequents = apriori_data['Consequent'].unique()

# Display unique consequents
unique_consequents


import matplotlib.pyplot as plt
import seaborn as sns

# Step 1: Aggregate data
consequents_data = apriori_data.groupby('Consequent').agg(
    Total_Lift=('RuleLift', 'sum'),
    Average_Confidence=('RuleConfidence%', 'mean')
).reset_index()

# Step 2: Sort by Total Lift for better visualization
consequents_data = consequents_data.sort_values(by='Total_Lift', ascending=False)

# Step 3: Create a stacked bar chart
plt.figure(figsize=(12, 8))
sns.barplot(x='Consequent', y='Total_Lift', data=consequents_data, color='b')
plt.xlabel('Consequents')
plt.ylabel('Total Lift')
plt.title('Total Lift Contribution by Consequent')
plt.xticks(rotation=45)
plt.show()

# Optionally, show the table as well for numerical reference
consequents_data



import pandas as pd

# Assuming you have your Apriori data in a DataFrame named 'apriori_data'
# This DataFrame should contain your Apriori analysis results with the necessary columns

n = 25

# Select and sort the top n rules by lift, for example
top_rules = apriori_data.nlargest(n, 'RuleLift')

# Create a stylized table with Pandas DataFrame
styled_table = top_rules.style.background_gradient(cmap='viridis', subset=['RuleLift']).set_properties(**{'text-align': 'left'})

# To display this table as an image or in a Jupyter notebook, you can simply output 'styled_table'
styled_table









import numpy as np

# Step 1: Aggregate data for all three metrics
consequents_data = apriori_data.groupby('Consequent').agg(
    Total_Support=('ItemSetSupport', 'sum'),
    Average_Confidence=('RuleConfidence%', 'mean'),  # Ensure that your confidence column matches the column name in your data
    Total_Lift=('RuleLift', 'sum')
).reset_index()

# Step 2: Sort by Total Lift for better visualization
consequents_data = consequents_data.sort_values(by='Total_Lift', ascending=False)

# Create a new figure for the grouped bar chart
plt.figure(figsize=(15, 10))

# Set the positions of the bars on the x-axis
r = np.arange(len(consequents_data))

# Width of the bars
barWidth = 0.25

# Plot data
plt.bar(r, consequents_data['Total_Support'], color='b', width=barWidth, edgecolor='grey', label='Total Support')
plt.bar(r + barWidth, consequents_data['Average_Confidence'], color='r', width=barWidth, edgecolor='grey', label='Average Confidence')
plt.bar(r + barWidth * 2, consequents_data['Total_Lift'], color='g', width=barWidth, edgecolor='grey', label='Total Lift')

# Add xticks on the middle of the group bars
plt.xlabel('Consequents', fontweight='bold')
plt.xticks([r + barWidth for r in range(len(consequents_data))], consequents_data['Consequent'], rotation=90)
plt.ylabel('Metric Values')
plt.title('Support, Confidence, and Lift for Each Consequent')

# Create legend & Show graphic
plt.legend()
plt.show()


