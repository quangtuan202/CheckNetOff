# Doing
import pyxlsb
import pandas as pd
import os
import datetime as dt
import xlsxwriter
import numpy as np

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
       'Dư đầu kỳ VND', 'Hạn thanh toán ','Phát sinh trong kỳ nguyên tệ',
       'Phát sinh trong kỳ > VND', 'Ngày thanh toán',
       'Thanh toán trong kỳ Nguyên tệ', 'Thanh toán trong kỳ > VND',
       'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND','Số ngày quá hạn thanh toán','Số chứng từ']
    df['TaiKhoan']=name
    dfList.append(df) 
dfFinal=pd.concat(dfList,axis=0)
dfFinal['TenTK']=dfFinal['TaiKhoan'].map(accountDict)
dfFinal=dfFinal.loc[(dfFinal['Số chứng từ'].notna())]
dfFinal['Count']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['Dư cuối kỳ > VND'].transform('count')

dfFinal['Val']=np.where(dfFinal['Dư cuối kỳ > VND']>0,1,-1)
dfFinal['AbsDuCuoiKy']=np.where(dfFinal['Dư cuối kỳ > VND']>0,dfFinal['Dư cuối kỳ > VND'],-dfFinal['Dư cuối kỳ > VND']).astype('float64')
dfFinal['MeanAbsDuCuoiKy']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['AbsDuCuoiKy'].transform('mean')
dfFinal['SumDuCuoiKy']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['Dư cuối kỳ > VND'].transform('sum')
dfFinal['SumDuCuoiKy_NT']=dfFinal.groupby(['Số đơn / EN ','Số kỳ','TaiKhoan'],as_index=False)['Dư cuối kỳ Nguyên tệ '].transform('sum').abs()
dfFinal['Percentage']=dfFinal['SumDuCuoiKy'].abs()/dfFinal['MeanAbsDuCuoiKy']
#
dfFinal['CLTG']=np.where(dfFinal['Percentage']<0.015,'CLTG','DCK')

#dfFinal['Row']=dfFinal.groupby(['Số đơn / EN '],as_index=False)['Count'].transform(lambda x: np.arange(1,len(x)+1))
# Ngay thanh toan la due date
dfFinal['Ngày thanh toán']=pd.to_datetime(dfFinal['Ngày thanh toán'],dayfirst=True)
dfFinal['Ngay so sanh']=pd.to_datetime(dt.datetime.now().date(),dayfirst=True)
dfFinal['Ngay qua han']=dfFinal['Ngày thanh toán']-dfFinal['Ngay so sanh']
dfFinal['Ngay qua han']=-dfFinal['Ngay qua han'].dt.days.astype('float64')
# Get correct name
dfReinsurerName=pd.read_excel("F:/QT - 2020/2020_REINSURANCE/+++ RI Outstanding/ReinsurerName.xlsx")
dfFinal=dfFinal.merge(dfReinsurerName,left_on='Mã KH',right_on='Code')
col=['Mã KH', 'Phòng', 'Reinsurer_Reinsured', 'Đối tượng bảo hiểm', 'Môi giới / Đại lý BH',
       'Số đơn / EN ', 'Số kỳ', 'Account month',
       'Số Ref (Đồng bảo\n hiểm nhận TBH)', 'Mã NT', 'Dư đầu kỳ nguyên tệ',
       'Dư đầu kỳ VND', 'Phát sinh trong kỳ nguyên tệ',
       'Phát sinh trong kỳ > VND', 'Ngày thanh toán',
       'Thanh toán trong kỳ Nguyên tệ', 'Thanh toán trong kỳ > VND',
       'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND','Ngay qua han', 'Số chứng từ', 'TaiKhoan', 'TenTK',
       'CLTG','SheetName']
dfFinal=dfFinal.loc[:,col]

#Apply format to date
def set_column_width(df,worksheet):
    maxLength = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
    for i, width in enumerate(maxLength):
        worksheet.set_column(i+1, i+1, width+1)

# Create worksheet for each name
filename=f"D:/CongNoPhaiTraTaiBaoHiem{dt.datetime.now().date()}.xlsx"
writer = pd.ExcelWriter(filename,engine='xlsxwriter')
workbook=writer.book

#Header format
header_format = workbook.add_format({'bold': True,
                                     'align': 'center',
                                     'valign': 'vcenter',
                                     'text_wrap': True,
                                     'fg_color': '#4295f5',
                                     'border': 1})
# Border format
border_fmt = workbook.add_format({'bottom':4, 'top':4, 'left':2, 'right':2})

# Create Df for each name and apply format, write to excel:
nameSet=list(set(dfFinal['SheetName']))
nameSet.sort()
dfList=[]
for name in nameSet:
    dfName=dfFinal.loc[dfFinal['SheetName']==name]
    dfList.append(dfName)

# Insert dfFinal
nameSet.insert(0,'1_All')
dfList.insert(0,dfFinal)
for i in range(len(dfList)):
    col_=['Mã KH', 'Phòng', 'Reinsurer_Reinsured', 'Đối tượng bảo hiểm', 'Môi giới / Đại lý BH',
       'Số đơn / EN ', 'Số kỳ', 'Account month',
       'Số Ref (Đồng bảo\n hiểm nhận TBH)', 'Mã NT', 'Dư đầu kỳ nguyên tệ',
       'Dư đầu kỳ VND', 'Phát sinh trong kỳ nguyên tệ',
       'Phát sinh trong kỳ > VND', 'Ngày thanh toán',
       'Thanh toán trong kỳ Nguyên tệ', 'Thanh toán trong kỳ > VND',
       'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND','Ngay qua han', 'Số chứng từ', 'TaiKhoan', 'TenTK',
       'CLTG']
    dfList[i]=dfList[i].loc[:,col_]
zipList=zip(dfList,nameSet)
for dfName,name in zipList:
    dfName.to_excel(writer, sheet_name=name, startrow=1, header=False)
    worksheet=writer.sheets[name]
    set_column_width(dfName,worksheet)  #Apply format to date
    # Apply header format
    for col_num, value in enumerate(dfName.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)
    worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(dfName), len(dfName.columns)), {'type': 'no_errors', 'format': border_fmt})
    worksheet.freeze_panes(1, 0) #Freeze pane
    worksheet.set_row(0, 30)
writer.save() # save 
