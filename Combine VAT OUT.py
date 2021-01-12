import pandas as pd
import pyxlsb
import os
import datetime as dt
path='D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/TAX/2020 - Tax/HTKK/Full'
df_list=[]

def taxRate(x):
    if x<lst[1]:
        return 'Khong chiu thue'
    elif x<lst[2]:
        return '0pt'
    elif x<lst[3]:
        return '5pt'
    elif x<lst[4]:
        return '10pt'
    else:
        return 'Khong ke khai'
for filename in os.listdir(path):
    if filename.endswith('xlsb'):
        df=pd.read_excel(f'{path}/{filename}',sheet_name='KeKhai',skiprows=16,header=None,engine='pyxlsb')
    else:
        df=pd.read_excel(f'{path}/{filename}',sheet_name='KeKhai',skiprows=16,header=None)
    df=df.iloc[:,1:12]
    df.columns=['STT','KyHieuMau','KyHieuHD','SoHD','NgayThang','KhachHang','MST','Hang','DoanhThu','VAT','Policy']
    a=list(df.STT.values)
    lst=[]
    for i in range(len(a)):
        if str(a[i]).startswith(('1.','2.','3.','4.','5.')):
            lst.append(i)
    df['ThueSuat']=df.index.map(taxRate)
    df['Thang']=filename[11:13]
    df=df.loc[df['Policy'].notna()]
    df_list.append(df)
dfFinal=pd.concat(df_list)
dfFinal=dfFinal.loc[dfFinal.KhachHang!='Ten']
dfFinal.DoanhThu=dfFinal.DoanhThu.fillna(0).astype('float64')
dfFinal.VAT=dfFinal.VAT.fillna(0).astype('float64')
time=str(dt.datetime.now()).replace(':','.')
dfFinal.to_excel(f'D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/TAX/2020 - Tax/12m{time}.xlsx')
import subprocess
subprocess.Popen([f'D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/TAX/2020 - Tax/12m{time}.xlsx'],shell=True)