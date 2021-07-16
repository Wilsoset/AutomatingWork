import requests
from bs4 import BeautifulSoup
import random
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

#Convert DF to a tolist
product = df['Name'].values.tolist()
print(product)
df['info'] = ""
df2 = pd.DataFrame(columns=["realname"])

#Loop through and search all the companies in the list
for name in product:
    A = ("Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36",
           "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.1 Safari/537.36",
           "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2227.0 Safari/537.36",
           )
    Agent = A[random.randrange(len(A))]
    headers = {'user-agent': Agent}
    url = 'https://google.com/search?q=' + name
    r = requests.get(url, headers=headers)
    content = BeautifulSoup(r.text, 'lxml')

    title = [title.text for title in content.find_all('h3')]
    description = [descr.text for descr in content.find_all('span')]

    df2 = df2.append({
    "realname":name,
    'title': title,
    'description': description
    },ignore_index=True)

df2['Crypto'] = df2['title'].astype(str).str.lower().str.contains('crypto') | df2['description'].astype(str).str.lower().str.contains('crypto') | df2['title'].astype(str).str.lower().str.contains('bitcoin') | df2['description'].astype(str).str.lower().str.contains('bitcoin')| df2['title'].astype(str).str.lower().str.contains('doge') | df2['description'].astype(str).str.lower().str.contains('doge')| df2['title'].astype(str).str.lower().str.contains('blockchain') | df2['description'].astype(str).str.lower().str.contains('blockchain')| df2['title'].astype(str).str.lower().str.contains('ethereum') | df2['description'].astype(str).str.lower().str.contains('ethereum')


#df[Crypto] = df['info'].str.lower().str.contains('test')
df2.to_excel('Mapped Crypto Export.xlsx', index=False)
print("Congrats, Companies Mapped!")
