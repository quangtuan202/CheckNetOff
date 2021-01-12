import pandas as pd
import numpy as np
import os

path="E:/OneDrive - khoavanhoc.edu.vn/cltg/congno/"
fx=pd.read_excel("E:/OneDrive - khoavanhoc.edu.vn/cltg/fx/exrate.xlsx",index_col='Ma')
fxBuy=fx.loc[:,['Buy']]
fxBuy=fxBuy.to_dict()['Buy']
fxSell=fx.loc[:,['Sell']]
fxSell=fxSell.to_dict()['Sell']
accName=pd.read_excel("E:/OneDrive - khoavanhoc.edu.vn/cltg/accname/accname.xlsx",index_col='Ma')
accName.index=accName.index.astype(str)
accName=accName.to_dict()['Ten']
dfList=[]
for filename in os.listdir(path):
    df=pd.read_excel(f'{path}/{filename}',header=None)
    account=df.at[4,1][14:20]
    df=df.iloc[7:]
    df.reset_index
    header=df.iloc[0]
    df=df[1:]
    df.columns=header
    #df=df.loc[(df['Mã NT']!='VND')&(df['Mã NT'].notna())]
    df=df.loc[df['Số chứng từ'].notna()]
    df['account']=account
    account=''
    dfList.append(df)
dfFinal=pd.concat(dfList)
dfFinal['1']=dfFinal['account'].str.slice(stop=1).astype(int)

#dfFinal['PhaiThu_PhaiTra']=pd.Series([dfFinal['1']==1,'PhaiThu','PhaiTra'))
dfFinal['PhaiThu_PhaiTra']=dfFinal['account'].apply(lambda x: 'PhaiThu' if x[0]=='1' else 'PhaiTra')
dfPhaiThu=dfFinal.loc[dfFinal['PhaiThu_PhaiTra']=='PhaiThu']
dfPhaiThu['TyGiaCuoiKy']=dfPhaiThu['Mã NT'].map(fxBuy)
dfPhaiThu['DuCkTyGiaCuoiKy']=dfPhaiThu['Dư cuối kỳ Nguyên tệ ']*dfPhaiThu['TyGiaCuoiKy']
dfPhaiThu['ChenhLechTyGiaCuoiKy']=dfPhaiThu['DuCkTyGiaCuoiKy']-dfPhaiThu['Dư cuối kỳ > VND']

dfPhaiTra=dfFinal.loc[dfFinal['PhaiThu_PhaiTra']=='PhaiTra']
dfPhaiTra['TyGiaCuoiKy']=dfPhaiTra['Mã NT'].map(fxSell)
dfPhaiTra['DuCkTyGiaCuoiKy']=dfPhaiTra['Dư cuối kỳ Nguyên tệ ']*dfPhaiTra['TyGiaCuoiKy']
dfPhaiTra['ChenhLechTyGiaCuoiKy']=-dfPhaiTra['DuCkTyGiaCuoiKy']+dfPhaiTra['Dư cuối kỳ > VND'] # doi dau so voi phai thu
dfFinal=pd.concat([dfPhaiThu,dfPhaiTra])
dfFinal['AccountName']=dfFinal['account'].map(accName)



