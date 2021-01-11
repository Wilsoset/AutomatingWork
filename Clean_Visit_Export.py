import os
import pandas as pd
import xlsxwriter

#Search for excel files in my desktop
path = os.getcwd()
files = os.listdir(path)
files_lsx = [f for f in files if f[-3:] == 'lsx']

#appened the excelfile to a Data Frame
df = pd.DataFrame()
for f in files_lsx:
    data = pd.read_excel(f)
    df = df.append(data)

#Helper Columns to find Test Accs
df['Concat'] = df['CO Email'].map(str) + df['Care Pro'].map(str) + df['CR First Name'].map(str) + df['CR Last Name'].map(str) + df['CO First Name'].map(str) + df['CO Last Name'].map(str)
df['Homage'] = df['Concat'].str.lower().str.contains('homage')
df['Test'] = df['Concat'].str.lower().str.contains('test')
df['Fake Operator'] = df['Operator Assigned'].str.contains('Janie Le|Joan Yin|Honey Mittal|Jane Le Hoang|Raymond Panganiban|Wei Sheng Poh|Rachel Ng |Alvin|Sohaib Muhammad|Adil|Usama Sadiq|Carl Dalid|Jia Yang Lai|Rui Ong|Manmeet Singh|Yu Jie Koh|Jeraldine',
na=False)
df['Fake Acc'] = df['Homage'].astype(str).str.lower().str.contains('true')  | df['Test'].astype(str).str.lower().str.contains('true') | df['Fake Operator'].astype(str).str.lower().str.contains('true')

#Highlight the rows where fake account = true
indexNames = df[df['Fake Acc'].astype(str).str.lower() == 'true' ].index
#Remove those rows
df.drop(indexNames,inplace=True)

#Remove Covid Orgs for B2B
df['Covid Org'] = df['B2B Organisation Name'].astype(str).str.contains('Agency of Integrated Care|Ministry of Health|Sheares Healthcare|Alvernia|Standard Chartered|Siemens|Health Promotion Board|Speedoc',
na = False)
#Highlight the rows where covid org = true
indexOrgs = df[df['Covid Org'].astype(str).str.lower() == 'true' ].index
#Remove those rows
df.drop(indexOrgs,inplace=True)

#Drop the Helper Columnhs so we can copy paste easy
df.drop(['Concat','Homage','Test','Fake Operator','Fake Acc','Covid Org'], axis=1, inplace=True)

#remove extra column created by Pandas
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

#Different Sheets for B2B and B2C
writer = pd.ExcelWriter('Cleaned Visit Export.xlsx', engine='xlsxwriter')
dfb2b = df.loc[df['CR Type'].str.contains('B2B')]
dfb2c = df.loc[df['CR Type'].str.contains('B2C')]
dfb2b.to_excel(writer, sheet_name='B2B', index = False)
dfb2c.to_excel(writer, sheet_name='B2C', index = False)
writer.save()

print("Congrats! Visit Export Cleaned!")
