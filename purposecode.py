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
df['ADVA'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('adva')
df['ALLW'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('allw')
df['BEXP'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('bexp')
df['BONU'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('bonu')
df['CFEE'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('cfee')
df['CHAR'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('char')
df['COMM'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('comm')
df['DIVD'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('divd')
df['EDUC'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('educ')
df['FEES'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('fees')
df['GDDS'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('gdds')
df['GDSV'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('gdsv')
df['HREC'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('hrec')
df['INTC'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('intc')
df['INVS'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('invs')
df['IVPT'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('ivpt')
df['LOAN'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('loan')
df['LOAR'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('loar')
df['MDCS'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('mdcs')
df['OTHR'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('othr')
df['PAYR'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('payr')
df['RELG'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('relg')
df['RENT'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('rent')
df['SAVG'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('savg')
df['SCVE'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('scve')
df['SERV'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('serv')
df['SUBS'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('subs')
df['SUPP'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('supp')
df['TAXS'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('taxs')
df['TPRP'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('tprp')
df['UBIL'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('ubil')
df['WX01'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('wx01')
df['WX04'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('wx04')
df['WX07'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('wx07')
df['WX08'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('wx08')
df['WX09'] = df['FIRST_3MONTHS_PURPOSE_CODES'].str.lower().str.contains('wx09')

df.to_excel('finance_cleaned.xlsx',index = False)

print("Congrats! Finance Files Cleaned!")
