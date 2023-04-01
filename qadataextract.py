import code
from email import header
import pandas as pd
import numpy as np
import os

def qadata(url_origi,round,folder,code,data_range,a):
    url= url_origi
    round = round
    topic = folder
    code =code
    url = url_origi[0:url_origi.find('edit')]+'export?format=xlsx'
    print(url)
    df = pd.read_excel(url,sheet_name = 0)
    
    df2 = df.rename({'Code อบรม':'code','ชื่อ-นามสกุล(English)เท่านั้น (ห้ามใช้ภาษาไทย)':'name','เบอร์โทร':'tel','ที่อยู่อีเมล':'email'},axis=1)#!น่าจะต้องแก้
    # print(df2.head(1))
    df2['code'] = df2['code'].str.upper()
    df2['code'] = df2['code'].str.replace(" ","")
    df2['code'] = df2['code'].str.find(code)
    df2 = df2.loc[df2['code']>=0]



    df2 = df2.loc[~df2['name'].duplicated()]
    # len(pd.unique(df2['name']))
    # df2['code'].value_counts()

    df2['no'] = df2.index
    
    df2 = df2.drop_duplicates('name').reset_index(drop=True)
    df2.index = df2.index+1
    df2.index.names = ['num']


    # with open('certnumlast.txt','r') as f:
    #     a = int(f.readline())
    # print(df2['no'])
    dd = df2.iloc[-1][-1]
    # print(pd.Series(np.arange(a,a+dd+2)))
    df2['cert_num'] = pd.Series(np.arange(a,a+dd+2))
    # print(df2.head())

    df3 = df2[['name','email','cert_num','tel']]
    lastcert= str(df2.iloc[-1][-1])
    # with open('certnumlast.txt','w') as f: #!อย่าลืม
    #     f.write(lastcert) #!อย่าลืม
    data_range=data_range
    df3.to_excel(f'{topic}/testdata{round}_all.xlsx',sheet_name='Sheet_name_1')
    count = 1
    if not os.path.isdir(topic):
        os.mkdir(topic)
    for i in range(0,len(df3),data_range):
        a=df3[i:i+data_range].reset_index(drop=True)
        a.index = a.index+1
        a.index.names = ['num']
        a.to_excel(f'{topic}/testdata{round}_{count}.xlsx',sheet_name='Sheet_name_1')
        count +=1


# qadata()