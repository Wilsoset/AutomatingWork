import os
import pandas as pd
import xlrd
cwd = os.path.abspath('C:\\Users\\Homage\\Desktop\\Python_Stuff\\')
files = os.listdir(cwd)

# Method 1 gets the first sheet of a given file
df = pd.DataFrame()
for file in files:
    if file.endswith('.xlsx'):
        df = df.append(pd.read_excel(file, sheet_name = 'visits' ), ignore_index=True)
        #remove extra column created by Pandas
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df.to_excel('compiled_sheets.xlsx',index = False)

print("Merge Completed!")
