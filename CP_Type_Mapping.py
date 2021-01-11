import os
import pandas as pd
import numpy as np

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
df['Concat'] = df['Email Address'].map(str) + df['CP Full Name'].map(str)
df['Homage'] = df['Concat'].str.lower().str.contains('homage')
df['Test'] = df['Concat'].str.lower().str.contains('test')
df['Fake Acc'] = df['Homage'].astype(str).str.lower().str.contains('true')  | df['Test'].astype(str).str.lower().str.contains('true')

#Mapping the Edited CP role Column wit the original Care Pro Type Column
df['CP Role - edited'] = df['Care Pro Type']
df['CP Role - edited'].fillna('Caregiver',inplace = True)
mapping = {'Therapist' : 'Therapist', 'Therapist - Occupational' : 'Therapist', 'Therapist - Speech' : 'Therapist', 'Therapist - Physio' : 'Therapist',
'Registered Nurse' : 'Nurse', 'Enrolled Nurse' : 'Nurse', 'Doctor' : 'Doctor', 'Swabber' : 'Swabber', 'Swab Assistant' : 'Swab Assistant',
'Non-Healthcare Student' : 'Caregiver', 'Healthcare Student' : 'Caregiver', 'Therapy Assistant' : 'Caregiver', 'Caregiver' : 'Caregiver',
 'Patient Care/Healthcare Assistant' : 'Caregiver', 'Patient Care or Healthcare Assistant' : 'Caregiver'}

df['CP Role - edited'].replace(mapping, inplace = True)

#Highlight the rows where fake account = true
indexNames = df[df['Fake Acc'].astype(str).str.lower() == 'true' ].index
#Remove those rows
df.drop(indexNames,inplace=True)

#Rearranging the Cp Role Edited Column from the far right to next to Care Pro type
cols = list(df.columns.values)
df = df [cols[0:6] + [cols[-1]] + cols[6:54] ]

df.to_excel('Mapped CP Export.xlsx', index=False)
print("Congrats, CP Role Mapped!")
