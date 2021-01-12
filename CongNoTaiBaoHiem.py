import pandas as pd 
import numpy as np
#import os
import datetime as dt


phaiThuGocFile='D:/Tai/131111.xls'
phaiTraTaiFile='D:/Tai/331311.xls'
dfPhaiThuGoc=pd.read_excel(phaiThuGocFile,skiprows=7)
dfPhaiTraTai=pd.read_excel(phaiTraTaiFile,skiprows=7)

dfPhaiThuGoc=dfPhaiThuGoc[(dfPhaiThuGoc['Mã NT'].notna())|(dfPhaiThuGoc['Số chứng từ'].notna())]
dfPhaiThuGoc['DueDate']=pd.to_datetime(dfPhaiThuGoc['Hạn thanh toán '],dayfirst=True)
dfPhaiThuGoc['Ngay den han']=dfPhaiThuGoc.groupby(['Số đơn / EN '])['DueDate'].transform('max')
dfPhaiThuGoc2=dfPhaiThuGoc.groupby(['Số đơn / EN ','TênKH','Ngay den han'],as_index=False)['Dư cuối kỳ > VND'].sum()
dfPhaiThuGoc2['Ngay so sanh']=pd.to_datetime(dt.datetime.now().date(),dayfirst=True)
dfPhaiThuGoc2['Ngay qua han']=dfPhaiThuGoc2['Ngay den han']-dfPhaiThuGoc2['Ngay so sanh']
dfPhaiThuGoc2['Ngay qua han']=-dfPhaiThuGoc2['Ngay qua han'].dt.days.astype('float64')

dfPhaiTraTai=dfPhaiTraTai[(dfPhaiTraTai['Mã NT'].notna())|(dfPhaiTraTai['Số chứng từ'].notna())]
dfPhaiTraTai['DueDate']=pd.to_datetime(dfPhaiTraTai['Hạn thanh toán '],dayfirst=True)
dfPhaiTraTai['Ngay den han']=dfPhaiTraTai.groupby(['Số đơn / EN '])['DueDate'].transform('max')
dfPhaiTraTai2=dfPhaiTraTai.groupby(['Số đơn / EN ','Mã KH','Số kỳ','Account month','TênKH','Ngay den han'],as_index=False)['Dư cuối kỳ Nguyên tệ ','Dư cuối kỳ > VND'].sum()
dfPhaiTraTai2['Ngay so sanh']=pd.to_datetime(dt.datetime.now().date(),dayfirst=True)
dfPhaiTraTai2['Ngay qua han']=dfPhaiTraTai2['Ngay den han']-dfPhaiTraTai2['Ngay so sanh']
dfPhaiTraTai2['Ngay qua han']=-dfPhaiTraTai2['Ngay qua han'].dt.days.astype('float64')

dfPhaiTraTai3=dfPhaiTraTai2.merge(dfPhaiThuGoc2,on='Số đơn / EN ',how='left',suffixes=(" Tái"," Gốc"))
dfPhaiTraTai3['Dư cuối kỳ > VND Gốc']=dfPhaiTraTai3['Dư cuối kỳ > VND Gốc'].fillna(0)
condition=dfPhaiTraTai3['Số đơn / EN '].str.contains('TTY|EXC|XOL')
dfPhaiTraTai3['FACT/TTY']=np.where(condition,'TTY','FACT')

dfPhaiTraTaiDuDieuKien=dfPhaiTraTai3.loc[(dfPhaiTraTai3['Ngay qua han Tái']>-30) & (dfPhaiTraTai3['Dư cuối kỳ > VND Gốc']==0)]

dfPhaiTraTaiChuaDuDieuKien=dfPhaiTraTai3.loc[(dfPhaiTraTai3['Ngay qua han Tái']>-30) & (dfPhaiTraTai3['Dư cuối kỳ > VND Gốc']!=0)]

writer = pd.ExcelWriter("D:/PhaiTraTai.xlsx",engine='xlsxwriter')
# Turn off the default header and skip one row to allow us to insert a
# user defined header.
dfPhaiTraTai3.to_excel(writer, sheet_name='Tong hop', startrow=1, header=False)
dfPhaiTraTaiDuDieuKien.to_excel(writer, sheet_name='CongNoTaiDuDieuKienTT', startrow=1, header=False)
dfPhaiTraTaiChuaDuDieuKien.to_excel(writer, sheet_name='CongNoTaiChuaDuDieuKienTT', startrow=1, header=False)

workbook  = writer.book
worksheet1 = writer.sheets['Tong hop']
worksheet2 = writer.sheets['CongNoTaiDuDieuKienTT']
worksheet3 = writer.sheets['CongNoTaiChuaDuDieuKienTT']

header_format = workbook.add_format({'bold': True,
                                     'align': 'center',
                                     'valign': 'vcenter',
                                     'text_wrap': True,
                                     'fg_color': '#4295f5',
                                     'border': 1})
for col_num, value in enumerate(dfPhaiTraTai3.columns.values):
    worksheet1.write(0, col_num + 1, value, header_format)

for col_num, value in enumerate(dfPhaiTraTaiDuDieuKien.columns.values):
    worksheet2.write(0, col_num + 1, value, header_format)

for col_num, value in enumerate(dfPhaiTraTaiChuaDuDieuKien.columns.values):
    worksheet3.write(0, col_num + 1, value, header_format)

worksheet1.freeze_panes(1, 0)
worksheet2.freeze_panes(1, 0)
worksheet3.freeze_panes(1, 0)

writer.save()