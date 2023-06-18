#!/usr/bin/env python
# coding: utf-8

# In[10]:


import pandas as pd

# Read the Excel file
df = pd.read_excel('Internship.xlsx')

# Example: Print the contents of the Excel file
print(df.head())


# In[14]:


import subprocess
import googlesearch

import openpyxl

# Load the Excel file
workbook = openpyxl.load_workbook('Internship.xlsx')

# Select the desired worksheet
worksheet = workbook['Sheet1']

# Define the column where you want to paste the result
column = 'B'  # Change it to the desired column letter

# Define the row where you want to paste the result
row = 2

# Get the values in the 'Company' column
companies = df['Company'].tolist()

# Iterate over each company
for company in companies:

    # Remove spaces
    company_name = company.replace(" ", "")
    query = f"{company_name}careers"
    
    
    search_results = googlesearch.search(query, num=1)
    
    # Extract the first result from the search results
    first_result = next(search_results, None)
    
    
    # Assign the result value to the corresponding cell
    cell = f'{column}{row}'
    worksheet[cell].hyperlink = first_result
    worksheet[cell].value = first_result
    row = row + 1
    
    
workbook.save('Final.xlsx')
    


# In[ ]:




