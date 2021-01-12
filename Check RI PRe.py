import pandas as pd
import os
import pyxlsb
df=pd.DataFrame
df_list=[]
use_col=['POLICY NUMBER','ENDORSEMENT NUMBER','BALANCE','REPOLICY TYPE']

path="D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/DOI CHIEU ACC-CPC/RI Pre-BK TTY/BK_RI"
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
func=lambda x, y: x if x!='NA' else y
df_merge['POLICY']=df_merge.apply(lambda x: func(x['POLICY NUMBER'],x['ENDORSEMENT NUMBER']),axis=1)
condition=df_merge['REPOLICY TYPE'].isin(['REIN132','REIN134','REIN175','REIN500','REIN741'])
df_merge=df_merge.loc[condition]
dfRI_groupby=df_merge.groupby(['POLICY'],as_index=False)['BALANCE'].sum()

#--------------------

df_tty=pd.read_excel("D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/DOI CHIEU ACC-CPC/RI Pre-BK TTY/TTY Pre Q3.xlsx",usecols=['Policy','Endorsement No','RI Premium'],sheet_name='BANG_KE')
df_tty['POLICY_TTY']=df_tty.apply(lambda x:func(x['Policy'],x['Endorsement No']),axis=1)
dfTty_groupby=df_tty.groupby(['POLICY_TTY'],as_index=False)['RI Premium'].sum()
df_check=dfTty_groupby.merge(dfRI_groupby,left_on='POLICY_TTY',right_on='POLICY',how='outer')
df_check['Diff']=df_check['RI Premium']-df_check['BALANCE']
df_check.to_csv("D:/Check_RI_Pre.csv")
