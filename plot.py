import pandas as pd
import numpy as np
import os
import urllib
import matplotlib.pyplot as plt
plt.style.use('seaborn-whitegrid')

df = pd.pandas.read_excel('C:\\Users\\Homage\\Desktop\\Python_Stuff\\Matrix Age and Income - Copy.xlsx')

df[["Seniors", "Median"]] = df[["Seniors", "Median"]].apply(pd.to_numeric)
df['Suburb'] = df['Suburb'].astype("|S")
df.info()

plt.scatter(df.Seniors,df.Median)
title('My plot title');
xlabel('Median');
ylabel('My y-axis title');
