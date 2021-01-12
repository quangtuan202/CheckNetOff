import pandas as pd 
import datetime as dt
path="C:/Users/tuandq/Downloads/bc_cn_tk_20201023093002.XLS"
df=pd.read_excel(path,skiprows=7)
df=df[(df['Mã NT'].notna())|(df['Số chứng từ'].notna())]
df['Date']=pd.to_datetime(df['Hạn thanh toán '],dayfirst=True)
df['Ngay thanh toan']=df.groupby(['Số đơn / EN '])['Date'].transform('max')
df_final=df.groupby(['Số đơn / EN ','Ngay thanh toan'],as_index=False)['Dư cuối kỳ > VND'].sum()
df_final['Ngay so sanh']=pd.to_datetime('30/09/2020',dayfirst=True)
df_final['Ngay qua han']=df_final['Ngay thanh toan']-df_final['Ngay so sanh']
df_final['Ngay qua han']=-df_final['Ngay qua han'].dt.days.astype('float64')


def provision(amount,day):
    if day<180:
        return 0
    elif day<360:
        return amount*0.3
    elif day<720:
        return amount*0.5
    elif day<1080:
        return amount*0.7
    else:
        return amount

def classify(day):
    if day<180:
        return '1. Phải thu phí bảo hiểm dưới 90 ngày'
    elif day<360:
        return '2. Phải thu phí bảo hiểm quá hạn từ 90 ngày đến dưới 01 năm'
    elif day<720:
        return '3. Phải thu phí bảo hiểm quá hạn từ 01 năm đến dưới 02 năm'
    else:
        return '4. Phải thu phí bảo hiểm quá hạn từ 02 năm trở lên'


df_final['Du phong']=df_final.apply(lambda x: provision(x['Dư cuối kỳ > VND'],x['Ngay qua han']),axis=1)
df_final['Phan loai']=df_final['Ngay qua han'].apply(classify)
df_final=df_final.groupby(['Phan loai'])['Dư cuối kỳ > VND','Du phong'].sum()