import os
import pandas as pd
import xlsxwriter

#Search for the specific file
df = pd.read_excel('C:/Users/Homage/Desktop/compiled_sheets.xlsx')

#Helper Columns to find Test Accs
df['Concat'] = df['CO Email'].map(str) + df['Care Pro'].map(str) + df['CR First Name'].map(str) + df['CR Last Name'].map(str) + df['CO First Name'].map(str) + df['CO Last Name'].map(str)
df['Homage'] = df['Concat'].str.lower().str.contains('homage')
df['Test'] = df['Concat'].str.lower().str.contains('test')
df['Total']=df['Care Pro'].str.lower().str.contains('total')
df['Fake Operator'] = df['Operator Assigned'].str.contains('Janie Le|Joan Yin|Honey Mittal|Jane Le Hoang|Raymond Panganiban|Wei Sheng Poh|Rachel Ng |Alvin|Sohaib Muhammad|Adil|Usama Sadiq|Carl Dalid|Jia Yang Lai|Rui Ong|Manmeet Singh|Yu Jie Koh|Jeraldine',
na=False)
df['Fake Acc'] = df['Homage'].astype(str).str.lower().str.contains('true')  | df['Test'].astype(str).str.lower().str.contains('true') | df['Fake Operator'].astype(str).str.lower().str.contains('true') | df['Total'].astype(str).str.lower().str.contains('true')

#Highlight the rows where fake account = true
indexNames = df[df['Fake Acc'].astype(str).str.lower() == 'true' ].index
#Remove those rows
df.drop(indexNames,inplace=True)

#Drop the Helper Columnhs so we can copy paste easy
df.drop(['Concat','Homage','Test','Fake Operator','Fake Acc','Total'], axis=1, inplace=True)

#remove extra column created by Pandas
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

df.to_excel('finance_cleaned.xlsx',index = False)

print("Congrats! Finance Files Cleaned!")
