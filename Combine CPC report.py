from numpy.lib.npyio import savez
import pandas as pd 
import subprocess
import os
import datetime

folder="D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/SQL_data/2020/202010/BC_ACC"
dfList=[]
for filename in os.listdir(folder):
    dfDict=pd.read_excel(f'{folder}/{filename}',sheet_name=None,header=None)
    lst=list(dfDict.values())
    new_header=lst[0].iloc[6,]
    lst[0]=lst[0].iloc[7:,]
    for i in range(len(lst)):
        lst[i].columns=new_header
    df=pd.concat(lst,axis=0)
    df=df.loc[:,df.columns[1:]]
    dfList.append(df)
dfFinal=pd.concat(dfList)
dfFinal=dfFinal.loc[:,dfFinal.columns[1:]]
save=f"D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/SQL_data/2020/202010/CTGS1-10{datetime.datetime.now()}.xlsx"
dfFinal.to_excel(save,index=False)
subprocess.Popen([save],shell=True)