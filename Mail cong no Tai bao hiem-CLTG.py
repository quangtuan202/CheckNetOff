import pyxlsb
import pandas as pd
import os

accountDict={'131211':'Phải thu phí nhận Tái Bảo hiểm',
            '131331':'Phải thu về hoàn phí nhượng tái bảo hiểm',
            '131411':'Phải thu bồi thường nhượng tái BH',
            '331311':'Phải trả phí nhượng Tái Bảo hiểm',
            '331411':'Phải trả bồi thường nhận Tái BH',
            '331431':'Phải trả về hoàn phí nhận Tái bảo hiểm'}

folder='F:/QT - 2020/2020_REINSURANCE/+++ RI Outstanding/OS_Nov'
dfList=[]
for filename in os.listdir(folder):
    name, ext = os.path.splitext(filename)
    df=pd.read_excel(f'{folder}/{filename}',sheet_name=None,header=None,engine='pyxlsb')
    df=list(df.values())
    df[0]=df[0].iloc[8:,:]
    df=pd.concat(df,axis=0)
    df=df.iloc[:,1:23]
    df.columns=['Mã KH', 'Phòng', 'TênKH', 'Đối tượng bảo hiểm',
       'Môi giới / Đại lý BH', 'Số đơn / EN ', 'Số kỳ', 'Account month',
       'Số Ref (Đồng bảo\n hiểm nhận TBH)', 'Mã NT', 'Dư đầu kỳ nguyên tệ',
       'Dư đầu kỳ VND', 'Hạn thanh toán ', 'Phát sinh trong kỳ nguyên tệ',
       'Phát sinh trong kỳ > VND', 'Ngày thanh toán',
       'Thanh toán trong kỳ Nguyên tệ', 'Thanh toán trong kỳ > VND',
       'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND',
       'Số ngày quá hạn thanh toán', 'Số chứng từ']
    df['TaiKhoan']=name
    dfList.append(df) 
dfFinal=pd.concat(dfList,axis=0)
dfFinal['TenTK']=dfFinal['TaiKhoan'].map(accountDict)
dfFinal=dfFinal.loc[(dfFinal['Số chứng từ'].notna())]
dfFinal['Count']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['Dư cuối kỳ > VND'].transform('count')
import numpy as np
dfFinal['Val']=np.where(dfFinal['Dư cuối kỳ > VND']>0,1,-1)
dfFinal['AbsDuCuoiKy']=np.where(dfFinal['Dư cuối kỳ > VND']>0,dfFinal['Dư cuối kỳ > VND'],-dfFinal['Dư cuối kỳ > VND']).astype('float64')
dfFinal['MeanAbsDuCuoiKy']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['AbsDuCuoiKy'].transform('mean')
dfFinal['SumDuCuoiKy']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['Dư cuối kỳ > VND'].transform('sum')
dfFinal['SumDuCuoiKy_NT']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['Dư cuối kỳ Nguyên tệ '].transform('sum').abs()
dfFinal['Percentage']=dfFinal['SumDuCuoiKy'].abs()/dfFinal['MeanAbsDuCuoiKy']
#
dfFinal['CLTG']=np.where(dfFinal['Percentage']<0.015,'CLTG','DCK')
#
dfFinal['Row']=dfFinal.groupby(['Số đơn / EN '],as_index=False)['Count'].transform(lambda x: np.arange(1,len(x)+1))