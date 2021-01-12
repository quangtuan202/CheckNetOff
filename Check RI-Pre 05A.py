import pandas as pd
import os
import pyxlsb
import subprocess
df=pd.DataFrame
df_list=[]
use_col=['POLICY NUMBER','ENDORSEMENT NUMBER','BALANCE','REPOLICY TYPE']

path="D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/DOI CHIEU ACC-CPC/RI Pre-BK TTY/23102020/RIPre1-9"
for filename in os.listdir(path):
    name, ext = os.path.splitext(filename)
    if ext=='.xls' or ext=='.xlsx':
        df=pd.read_excel(f'{path}/{filename}',usecols=use_col,skiprows=2)
        df_list.append(df)
    elif ext=='.xlsb':
        df=pd.read_excel(f'{path}/{filename}',usecols=use_col,skiprows=2,engine='pyxlsb')
        df_list.append(df)
df_merge=pd.concat(df_list)
df_merge['ENDORSEMENT NUMBER']=df_merge['ENDORSEMENT NUMBER'].fillna('NA')
func=lambda x, y: x if y!='NA' else x
df_merge['POLICY']=df_merge.apply(lambda x: func(x['POLICY NUMBER'],x['ENDORSEMENT NUMBER']),axis=1)
#condition=df_merge['REPOLICY TYPE'].isin(['REIN132','REIN134','REIN175','REIN500','REIN741']) #-TTY
condition=df_merge['REPOLICY TYPE'].isin(['003','REIN138']) #-FACT
df_merge=df_merge.loc[condition]
dfRI_groupby=df_merge.groupby(['POLICY'],as_index=False)['BALANCE'].sum()

#--------------------
path2="D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/DOI CHIEU ACC-CPC/RI Pre-BK TTY/23102020/05A1-9"
df_list2=[]
use_col2=['Số đơn/ Số Endor','Tỷ giá','Phí nhượng TBH FAC']
for filename in os.listdir(path2):
    name, ext = os.path.splitext(filename)
    if ext=='.xls' or ext=='.xlsx':
        df=pd.read_excel(f'{path2}/{filename}',usecols=use_col2,skiprows=7)
        df_list2.append(df)
    elif ext=='.xlsb':
        df=pd.read_excel(f'{path2}/{filename}',usecols=use_col2,skiprows=7,engine='pyxlsb')
        df_list2.append(df)
df_merge_5A=pd.concat(df_list2)
df_merge_5A=df_merge_5A.loc[(df_merge_5A['Phí nhượng TBH FAC'].notna()) & (df_merge_5A['Phí nhượng TBH FAC']!=0)]
df_merge_5A[['Tỷ giá','Phí nhượng TBH FAC']]=df_merge_5A[['Tỷ giá','Phí nhượng TBH FAC']].fillna(0)
#df_merge_5A=df_merge_5A.loc[df_merge_5A['Phí nhượng TBH FAC']!=0]
func=lambda x,y: x/y if y!=0 else 0
df_merge_5A['RI Fact NT']=df_merge_5A.apply(lambda x: func(x['Phí nhượng TBH FAC'],x['Tỷ giá']),axis=1)
df_merge_5A=df_merge_5A.groupby(['Số đơn/ Số Endor'],as_index=False)['RI Fact NT'].sum()
#---------------------------------------------------------------------------------------------
df_check=df_merge_5A.merge(dfRI_groupby,left_on='Số đơn/ Số Endor',right_on='POLICY',how='outer')
df_check[['Số đơn/ Số Endor','POLICY']]=df_check[['Số đơn/ Số Endor','POLICY']].fillna('NA')
df_check[['RI Fact NT','BALANCE']]=df_check[['RI Fact NT','BALANCE']].fillna(0)
df_check['Diff']=df_check['RI Fact NT']-df_check['BALANCE']
df_check.to_csv("D:/Check_RI_Pre.csv")
subprocess.Popen(["D:/Check_RI_Pre.csv"],shell=True)

