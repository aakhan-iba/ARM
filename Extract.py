#!/usr/bin/env python3
# Two Risk


#from shutil import copyfile
#from openpyxl import drawing
import random
import os.path
import openpyxl
import xlsxwriter
paths="D:/zipC/ioT"
#paths2="C:/Users/ladde/Desktop/zipC/filesNew"
#names=os.listdir(paths)

#print(names)

Source=paths+"/6data.xlsx"
#Source2=paths+"/6_1_data.xlsx"
Target=paths+"/7_1_Table.xlsx"


import pandas as pd

# Read the Excel file into a Pandas DataFrame
df = pd.read_excel(Source)
##
##df.to_excel(Source2, index=False)
##df = pd.read_excel(Source2)


N=len(df)
# Concatenate the columns into a new column named 'Concatenated_Column'
#df['Concatenated_Column'] = df['Fitness Level'] + ', ' + df['Physical Activity'] + ', ' + df['Blood Pressure'] + ', ' + df['Heart Beat']

# If you want to include column headers in the concatenated result
df['Rule'] = df['Fitness Level'] + ', ' + df['Physical Activity'] + ', BP ' + df['Blood Pressure'] + ', HR ' + df['Heart Beat']
df['String2'] = df['Fitness Level'] + ', ' + df['Physical Activity'] + ', BP ' + df['Blood Pressure'] + ', HR ' + df['Heart Beat'] + ' -> ' + df['Risk of Occurance']

# Save the DataFrame back to the Excel file
#df.to_excel(Out, index=False)


# Assuming your DataFrame is named 'df'
# Replace 'output_file.xlsx' with the path to your file if needed
#df = pd.read_excel(Out)

unique_counts = df['Rule'].value_counts().reset_index()
unique_counts.columns = ['Rule', 'String Count']

# Count occurrences of "YES" and "NO" for each unique string
df_yes = df[df['Risk of Occurance'] == 'Yes']
df_no = df[df['Risk of Occurance'] == 'No']

count_yes = df_yes['Rule'].value_counts().reset_index()
count_yes.columns = ['Rule', 'Frequency']

count_no = df_no['Rule'].value_counts().reset_index()
count_no.columns = ['Rule', 'Count_NO']

# Merge DataFrames on 'Unique_String'
final_output = pd.merge(unique_counts, count_yes, on='Rule', how='left').fillna(0)
#final_output = pd.merge(final_output, count_no, on='Rule', how='left').fillna(0)

# Save the result to a new Excel file
#final_output.to_excel('D:/zipc/iot/final_output_with_counts.xlsx', index=False)


#final_output = pd.read_excel('D:/zipc/iot/final_output_with_counts.xlsx')
# Calculate Support, Confidence, and Lift

final_output['Support'] = final_output['Frequency'] / N
final_output['Confidence (%)'] = (final_output['Frequency'] / final_output['String Count'])*100  #Confidence_YES
final_output['Lift'] = N*final_output['Frequency'] / ((final_output['String Count']) * len(df_yes))  #Lift_YES

# Concatenate "Yes" with Rule
final_output['Rule'] = final_output['Rule'] +' -> Yes'

# Insert a new column 'Serial No.' at the leftmost position with serial numbers
final_output.insert(0, 'S.No.', range(1, len(final_output) + 1))


# Save the result to a new Excel file
final_output.to_excel(Target, index=False)

