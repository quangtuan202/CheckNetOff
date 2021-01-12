import pandas as pd 
import os
path="D:/131111"
dfDict=[]
for filename in os.listdir(path):
    df=pd.read_excel(path+"/"+filename,sheet_name=None,header=None)
    dfDict.append(df) # return a dictionary
dfListFinal=[]
for i in range(len(dfDict)):
    dfList=list(dfDict[i].values())
    dfList[0]=dfList[0].iloc[5:,:]
    df=pd.concat(dfList,axis=0)
    dfListFinal.append(df)
df=pd.concat(dfListFinal,axis=0)
df=df.iloc[:,1:]
df.columns=['Loại chứng từ','Số chứng từ','Ngày chứng từ','Số đơn bảo hiểm','Số đơn chứng từ','Số đơn đối trừ của dòng hạch toán','Mã khách hàng nợ ','Tên khách hàng nợ ','Mã khách hàng có','Tên khách hàng có','Nội dung','TK đối ứng_Bên nợ','TK chi tiết_Bên nợ','Số tiền_Bên nợ','Số tiền VND_Bên nợ','TK đối ứng_Bên có','TK chi tiet_Bên có','Số tiền_Bên có','Số tiền VND_Bên có','Mã NT','Mã PC','Ghi chú']

df[['Số tiền_Bên nợ','Số tiền VND_Bên nợ','Số tiền_Bên có','Số tiền VND_Bên có']]=df[['Số tiền_Bên nợ','Số tiền VND_Bên nợ','Số tiền_Bên có','Số tiền VND_Bên có']].astype('float64')

df=df.loc[(df['Loại chứng từ']!='SỐ DƯ CUỐI KỲ NGUYÊN TỆ') & (df['Loại chứng từ']!='Tổng cộng:')]

df.to_hdf('d:/131111_2016_2019.h5', key='df', mode='w',index=False)