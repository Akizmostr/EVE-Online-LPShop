import requests
import lxml.html as lh
import pandas as pd
import os
import openpyxl as xl
import itertools
import os.path

url = 'https://www.fuzzwork.co.uk/lpstore/sell/10000002/1000135'

page = requests.get(url)

doc = lh.fromstring(page.content)

tr_elements = doc.xpath('//tr')

tr_elements = list(filter(lambda T: len(T)==10,tr_elements))
col=[]
i=0

for t in tr_elements[0]:
    i+=1
    name=t.text_content()
    col.append((name,[]))
for j in range(1, len(tr_elements)):
    T = tr_elements[j]

    i = 0

    for t in T.iterchildren():
        data = t.text_content()

        if i > 0:
            try:
                data = data.replace(',','')
                data = float(data)
                data = round(data, 0)
                data = int(data)
            except:
                pass
        col[i][1].append(data)

        i += 1
dict={title:column for (title,column) in col}
df=pd.DataFrame(dict)
df.drop(df.columns[[0, 2, 4, 5, 6, 7, 8]], axis = 1, inplace = True)

# selecting rows based on condition
snake_df = df[df['Item'].str.contains('Snake')]
asklep_df = df[df['Item'].str.contains('Asklepian')]
frames = [asklep_df, snake_df]
df = pd.concat(frames)
df = df.reset_index(drop=True)

print(df)

if not os.path.isfile('C:/Users/1/Desktop/lp.xlsm'):
    df.to_excel('C:/Users/1/Desktop/lp.xlsm', index=False)

workbook = xl.load_workbook(filename='C:/Users/1/Desktop/lp.xlsm', keep_vba=True, keep_links=True)

sheet = workbook.active

# for index, row in df.iterrows():
#     print(index, row['Item'])

# for (xl_row, (index, pd_row)) in zip(sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=3, max_col=3), df.iterrows()):
#     xl_row[0].value = pd_row['isk/lp']

max_row = len(df.index) + 1

for (index, pd_row) in df.iterrows():
    item = pd_row['Item']
    isk_lp = pd_row['isk/lp']
    for xl_row in sheet.iter_rows(min_row=2, max_row=max_row, min_col = 3, max_col=5):
        if xl_row[0].value == item:
            delta = isk_lp - xl_row[1].value
            xl_row[1].value = isk_lp
            xl_row[2].value = delta
            break


workbook.save(filename=('C:/Users/1/Desktop/lp.xlsm'))
os.startfile('C:/Users/1/Desktop/lp.xlsm')