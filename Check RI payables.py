import pandas as pd 
import numpy as np
import os
import datetime as dt
import xlsxwriter


Data1311111NamTruoc=r'D:/131111_2016_2019.h5' #HDF file for previous year
Data1311111NamNay=r'F:/DATA/RI_PAYABLE/131111_NAM_NAY/131111.XLS' #File so chi tiet 131111 current year
phaiThuGocFile='D:/Tai/131111.xls' # File so du cong no phi goc
phaiTraTaiFile='D:/Tai/331311.xls' # File so du cong no phi tai
customerNameLastYearsFile=r'F:/DATA/RI_PAYABLE/BC05A_NAM_TRUOC/TenKhachHang.h5'
customerNameThisYearsFolder=r'F:/DATA/RI_PAYABLE/BC05A_NAM_NAY'
dfPhaiThuGoc=pd.read_excel(phaiThuGocFile,skiprows=7)
dfPhaiTraTai=pd.read_excel(phaiTraTaiFile,skiprows=7)

# Transform data of GWP receivales
dfPhaiThuGoc=dfPhaiThuGoc.loc[(dfPhaiThuGoc['Số chứng từ'].notna())]
dfPhaiThuGoc[['Mã KH','TênKH']]=dfPhaiThuGoc[['Mã KH','TênKH']].fillna('N/A')
dfPhaiThuGoc['DueDate']=pd.to_datetime(dfPhaiThuGoc['Hạn thanh toán '],dayfirst=True)
dfPhaiThuGoc['Ngay den han']=dfPhaiThuGoc.groupby(['Số đơn / EN '])[['DueDate']].transform('max')
dfPhaiThuGoc2=dfPhaiThuGoc.groupby(['Số đơn / EN ','TênKH','Ngay den han'],as_index=False)[['Dư cuối kỳ > VND']].sum()
dfPhaiThuGoc2['Ngay so sanh']=pd.to_datetime(dt.datetime.now().date(),dayfirst=True)
dfPhaiThuGoc2['Ngay qua han']=dfPhaiThuGoc2['Ngay den han']-dfPhaiThuGoc2['Ngay so sanh']
dfPhaiThuGoc2['Ngay qua han']=-dfPhaiThuGoc2['Ngay qua han'].dt.days.astype('float64')

# Transform data of RI payales
dfPhaiTraTai=dfPhaiTraTai.loc[(dfPhaiTraTai['Số chứng từ'].notna())]
dfPhaiTraTai[['Mã KH','TênKH']]=dfPhaiTraTai[['Mã KH','TênKH']].fillna('N/A')
dfPhaiTraTai['DueDate']=pd.to_datetime(dfPhaiTraTai['Hạn thanh toán '],dayfirst=True)
dfPhaiTraTai['Ngay den han']=dfPhaiTraTai.groupby(['Số đơn / EN '])[['DueDate']].transform('max')
dfPhaiTraTai2=dfPhaiTraTai.groupby(['Số đơn / EN ','Mã KH','Số kỳ','Account month','TênKH','Ngay den han'],as_index=False)[['Dư cuối kỳ Nguyên tệ ','Dư cuối kỳ > VND']].sum()
dfPhaiTraTai2['Ngay so sanh']=pd.to_datetime(dt.datetime.now().date(),dayfirst=True)
dfPhaiTraTai2['Ngay qua han']=dfPhaiTraTai2['Ngay den han']-dfPhaiTraTai2['Ngay so sanh']
dfPhaiTraTai2['Ngay qua han']=-dfPhaiTraTai2['Ngay qua han'].dt.days.astype('float64')

# Merge
dfPhaiTraTai3=dfPhaiTraTai2.merge(dfPhaiThuGoc2,on='Số đơn / EN ',how='left',suffixes=(" Tái"," Gốc"))
dfPhaiTraTai3['Dư cuối kỳ > VND Gốc']=dfPhaiTraTai3['Dư cuối kỳ > VND Gốc'].fillna(0)
condition=dfPhaiTraTai3['Số đơn / EN '].str.contains('TTY|EXC|XOL')
dfPhaiTraTai3['FACT/TTY']=np.where(condition,'TTY','FACT')

# Create function for policy to dept
# Use np.select for performance
def policyDept(policy):
    condition=[
        policy.str.contains('SYCAR'),
        policy.str.startswith('0') & policy.str.endswith('T'),
        policy.str.contains('CTP'),
        policy.str.contains('HB'),
        policy.str.contains('HY'),
        policy.str.contains('HL'),
        policy.str.contains('SY'),
        policy.str.contains('SN'),
        policy.str.contains('HB'),
        policy.str.contains('SB'),  
        policy.str.contains('HP'),
        policy.str.contains('HR'),
        policy.str.contains('HU'),
        policy.str.contains('HG'),
        policy.str.contains('DR'),
        policy.str.contains('VR'),
        policy.str.contains('SG'),
        policy.str.contains('SR'),
        policy.str.contains('FR'),
        policy.str.contains('HN'),
        policy.str.contains('EN')]

    output=[
        policy.str.find('SYCAR')+5,
        policy.str.len()-1,
        policy.str.find('CTP')+3,
        policy.str.find('HB')+3,
        policy.str.find('HY')+3,
        policy.str.find('HL')+3,
        policy.str.find('SY')+3,
        policy.str.find('SN')+3,
        policy.str.find('HB')+3,
        policy.str.find('SB')+3,  
        policy.str.find('HP')+3,
        policy.str.find('HR')+3,
        policy.str.find('HU')+3,
        policy.str.find('HG')+3,
        policy.str.find('DR')+3,
        policy.str.find('VR')+3,
        policy.str.find('SG')+3,
        policy.str.find('SR')+3,
        policy.str.find('FR')+3,
        policy.str.find('HN')+3,
        policy.str.find('EN')]
    return [a[:b] for a, b in zip(policy,np.select(condition,output,policy.str.len()))]

#Read H5 file of previous year for customer name
dfCustomerNameLastYears=pd.read_hdf(customerNameLastYearsFile)

#Read excel file of BCTH05A
dfDict=[]
for filename in os.listdir(customerNameThisYearsFolder):
    df=pd.read_excel(f'{customerNameThisYearsFolder}/{filename}',sheet_name=None,header=None)
    dfDict.append(df) # return a dictionary
dfListFinal=[]
for i in range(len(dfDict)):
    dfList=list(dfDict[i].values())
    dfList[0]=dfList[0].iloc[8:,:]
    df=pd.concat(dfList,axis=0)
    dfListFinal.append(df)
dfCustomerNameThisYear=pd.concat(dfListFinal,axis=0)
dfCustomerNameThisYear=dfCustomerNameThisYear.iloc[:,[2,4,9,10]]
dfCustomerNameThisYear.columns=['NT_TT','Policy','TenKH','TenDoiTuongBH']
dfCustomerNameThisYear['SoDonMaPhong']=policyDept(dfCustomerNameThisYear['Policy'])
dfCustomerNameThisYear['KhachHang']=np.where(dfCustomerNameThisYear['TenDoiTuongBH'].notna(),dfCustomerNameThisYear['TenDoiTuongBH'],dfCustomerNameThisYear['TenKH'])
dfCustomerNameThisYear=dfCustomerNameThisYear.loc[:,['SoDonMaPhong','KhachHang']]

#Concatenate TenKH last year and this year
dfCustomerNameFull=pd.concat([dfCustomerNameLastYears,dfCustomerNameThisYear],axis=0)

# Remove duplicates
dfCustomerNameFull['RowNum']=dfCustomerNameFull.groupby(['SoDonMaPhong'])[['KhachHang']].transform(lambda x: np.arange(1, len(x)+1)) #Mimic Row_number() Over (Partition by) SQL
dfCustomerNameFull=dfCustomerNameFull.loc[dfCustomerNameFull.RowNum==1,['SoDonMaPhong','KhachHang']] # Select 1st row of each group

# Merge dfPhaiTraTai3 with dfCustomerNameFull to get customer name
dfPhaiTraTai3['SoDonMaPhong']=policyDept(dfPhaiTraTai3['Số đơn / EN '])
dfPhaiTraTai3=dfPhaiTraTai3.merge(dfCustomerNameFull,on='SoDonMaPhong',how='left')
dfPhaiTraTai3=dfPhaiTraTai3.loc[:,['Số đơn / EN ', 'Mã KH', 'Số kỳ', 'Account month', 'TênKH Tái',
       'Ngay den han Tái', 'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND Tái',
       'Ngay so sanh Tái', 'Ngay qua han Tái',  'KhachHang', 'Ngay den han Gốc',
       'Dư cuối kỳ > VND Gốc', 'Ngay so sanh Gốc', 'Ngay qua han Gốc',
       'FACT/TTY','TênKH Gốc']]

# Requirements satisfied
dfPhaiTraTaiDuDieuKien=dfPhaiTraTai3.loc[(dfPhaiTraTai3['Ngay qua han Tái']>-30) & (dfPhaiTraTai3['Dư cuối kỳ > VND Gốc']==0)]

# Requirements not satisfied
dfPhaiTraTaiChuaDuDieuKien=dfPhaiTraTai3.loc[(dfPhaiTraTai3['Ngay qua han Tái']>-30) & (dfPhaiTraTai3['Dư cuối kỳ > VND Gốc']!=0)]

# Not due
dfPhaiTraTaiChuaDenHan=dfPhaiTraTai3.loc[(dfPhaiTraTai3['Ngay qua han Tái']<=-30)]

# Read file HDF 2016-2019
df131111_2016_2019=pd.read_hdf(Data1311111NamTruoc)
df131111_2016_2019=df131111_2016_2019.loc[df131111_2016_2019['TK chi tiet_Bên có'].notna(),['Số chứng từ','Ngày chứng từ','Số đơn bảo hiểm','Số đơn chứng từ','TK chi tiet_Bên có','Số tiền_Bên có','Số tiền VND_Bên có']]

# Read file XLS 2020
df131111_2020=pd.read_excel(Data1311111NamNay,sheet_name=None,header=None)
dfList=list(df131111_2020.values())
dfList[0]=dfList[0].iloc[5:,:]
df131111_2020=pd.concat(dfList,axis=0)
df131111_2020=df131111_2020.iloc[:,1:]
df131111_2020.columns=['Loại chứng từ','Số chứng từ','Ngày chứng từ','Số đơn bảo hiểm','Số đơn chứng từ','Số đơn đối trừ của dòng hạch toán','Mã khách hàng nợ ','Tên khách hàng nợ ','Mã khách hàng có','Tên khách hàng có','Nội dung','TK đối ứng_Bên nợ','TK chi tiết_Bên nợ','Số tiền_Bên nợ','Số tiền VND_Bên nợ','TK đối ứng_Bên có','TK chi tiet_Bên có','Số tiền_Bên có','Số tiền VND_Bên có','Mã NT','Mã PC','Ghi chú']
df131111_2020[['Số tiền_Bên nợ','Số tiền VND_Bên nợ','Số tiền_Bên có','Số tiền VND_Bên có']]=df131111_2020[['Số tiền_Bên nợ','Số tiền VND_Bên nợ','Số tiền_Bên có','Số tiền VND_Bên có']].astype('float64')
df131111_2020=df131111_2020.loc[(df131111_2020['Loại chứng từ']!='SỐ DƯ CUỐI KỲ NGUYÊN TỆ') & (df131111_2020['Loại chứng từ']!='Tổng cộng:')]
df131111_2020=df131111_2020.loc[df131111_2020['TK chi tiet_Bên có'].notna(),['Số chứng từ','Ngày chứng từ','Số đơn bảo hiểm','Số đơn chứng từ','TK chi tiet_Bên có','Số tiền_Bên có','Số tiền VND_Bên có']]

# Combine DF_2016-2019 & DF_2020
df131111UpToDate=pd.concat([df131111_2016_2019,df131111_2020],axis=0)

# Lambda function to get correct policy
#getPolicy=lambda x,y: x if len(str(x))>5 else y 

# Apply function to DF_ALL
#df131111UpToDate['Policy']=df131111UpToDate.apply(lambda x: getPolicy(x['Số đơn bảo hiểm'],x['Số đơn chứng từ']),axis=1)
df131111UpToDate['Policy']=np.where(df131111UpToDate['Số đơn bảo hiểm'].str.len()>5,df131111UpToDate['Số đơn bảo hiểm'],df131111UpToDate['Số đơn chứng từ'])
df131111UpToDate['Policy']=df131111UpToDate['Policy'].fillna('N/A')
df131111UpToDate['Policy']=np.where(df131111UpToDate['Policy'].str.startswith('UPDATE_'),df131111UpToDate['Policy'].str.slice(start=7),df131111UpToDate['Policy'])

# Filter policy with Receivables of Zero balance 
df131111UpTpDate=df131111UpToDate.loc[df131111UpToDate['Policy'].isin(dfPhaiTraTaiDuDieuKien['Số đơn / EN '])]

# Create Paid info column
df131111UpToDate['Chứng từ thu phí gốc']=df131111UpToDate['Số chứng từ']+' : '+df131111UpToDate['Số tiền VND_Bên có'].astype('string')+' : '+df131111UpToDate['Ngày chứng từ']

#Concatenate info
df131111UpToDate=df131111UpToDate.groupby(['Policy'],as_index=False)['Chứng từ thu phí gốc'].apply('\n'.join).reset_index()

#Merge dfPhaiTraTaiDuDieuKien and df131111UpToDate
dfPhaiTraTaiDuDieuKien=dfPhaiTraTaiDuDieuKien.merge(df131111UpToDate,left_on='Số đơn / EN ',right_on='Policy',how='left').loc[:,['Số đơn / EN ', 'Mã KH', 'Số kỳ', 'Account month', 'TênKH Tái', 'Ngay den han Tái', 'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND Tái','Ngay so sanh Tái', 'Ngay qua han Tái', 'KhachHang', 'Ngay den han Gốc','Dư cuối kỳ > VND Gốc', 'Ngay so sanh Gốc', 'Ngay qua han Gốc','FACT/TTY','Chứng từ thu phí gốc']]

dfPhaiTraTaiChuaDuDieuKien=dfPhaiTraTaiChuaDuDieuKien.merge(df131111UpToDate,left_on='Số đơn / EN ',right_on='Policy',how='left').loc[:,['Số đơn / EN ', 'Mã KH', 'Số kỳ', 'Account month', 'TênKH Tái','Ngay den han Tái', 'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND Tái','Ngay so sanh Tái', 'Ngay qua han Tái', 'TênKH Gốc', 'Ngay den han Gốc','Dư cuối kỳ > VND Gốc', 'Ngay so sanh Gốc', 'Ngay qua han Gốc','FACT/TTY','Chứng từ thu phí gốc']]

dfPhaiTraTaiChuaDenHan=dfPhaiTraTaiChuaDenHan.merge(df131111UpToDate,left_on='Số đơn / EN ',right_on='Policy',how='left').loc[:,['Số đơn / EN ', 'Mã KH', 'Số kỳ', 'Account month', 'TênKH Tái','Ngay den han Tái', 'Dư cuối kỳ Nguyên tệ ', 'Dư cuối kỳ > VND Tái','Ngay so sanh Tái', 'Ngay qua han Tái', 'KhachHang', 'Ngay den han Gốc','Dư cuối kỳ > VND Gốc', 'Ngay so sanh Gốc', 'Ngay qua han Gốc','FACT/TTY','Chứng từ thu phí gốc']]

# Format datetime
for col in ['Ngay den han Tái','Ngay so sanh Tái','Ngay den han Gốc','Ngay so sanh Gốc']:
    dfPhaiTraTai3[col]=dfPhaiTraTai3[col].dt.strftime('%d/%m/%Y')
    dfPhaiTraTaiDuDieuKien[col]=dfPhaiTraTaiDuDieuKien[col].dt.strftime('%d/%m/%Y')
    dfPhaiTraTaiChuaDuDieuKien[col]=dfPhaiTraTaiChuaDuDieuKien[col].dt.strftime('%d/%m/%Y')
    dfPhaiTraTaiChuaDenHan[col]=dfPhaiTraTaiChuaDenHan[col].dt.strftime('%d/%m/%Y')

#Write DFs to Excel file and apply format
filename=f"D:/CongNoPhaiTraTaiBaoHiem{dt.datetime.now().date()}.xlsx"
writer = pd.ExcelWriter(filename,engine='xlsxwriter')
# Turn off the default header and skip one row to allow us to insert a
# User defined header.
dfPhaiTraTai3.to_excel(writer, sheet_name='Tong hop', startrow=1, header=False)
dfPhaiTraTaiDuDieuKien.to_excel(writer, sheet_name='CongNoTaiDuDieuKienTT', startrow=1, header=False)
dfPhaiTraTaiChuaDuDieuKien.to_excel(writer, sheet_name='CongNoTaiChuaDuDieuKienTT', startrow=1, header=False)
dfPhaiTraTaiChuaDenHan.to_excel(writer, sheet_name='CongNoTaiChuaDenHan', startrow=1, header=False)
workbook  = writer.book
worksheet1 = writer.sheets['Tong hop']
worksheet2 = writer.sheets['CongNoTaiDuDieuKienTT']
worksheet3 = writer.sheets['CongNoTaiChuaDuDieuKienTT']
worksheet4 = writer.sheets['CongNoTaiChuaDenHan']

#Apply format to number 
number_format = workbook.add_format({'num_format': '#,##0'})
worksheet1.set_column('H:I', None, number_format)
worksheet1.set_column('N:N', None, number_format)

worksheet2.set_column('H:I', None, number_format)
worksheet2.set_column('N:N', None, number_format)

worksheet3.set_column('H:I', None, number_format)
worksheet3.set_column('N:N', None, number_format)

worksheet4.set_column('H:I', None, number_format)
worksheet4.set_column('N:N', None, number_format)

#Apply format to date
def set_column_width(df,worksheet):
    maxLength = [max([len(str(s)) for s in df[col].values]) for col in df.columns]
    for i, width in enumerate(maxLength):
        worksheet.set_column(i+1, i+1, width+1)

set_column_width(dfPhaiTraTai3,worksheet1)    
set_column_width(dfPhaiTraTaiChuaDenHan,worksheet2)
set_column_width(dfPhaiTraTaiChuaDuDieuKien,worksheet3)   
set_column_width(dfPhaiTraTaiDuDieuKien,worksheet4)
    
#date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
#worksheet1.set_column('G:G', None, date_format)
#worksheet1.set_column('J:J', None, date_format)
#worksheet1.set_column('M:M', None, date_format)
#worksheet1.set_column('O:O', None, date_format)

#worksheet2.set_column('G:G', None, date_format)
#worksheet2.set_column('J:J', None, date_format)
#worksheet2.set_column('M:M', None, date_format)
#worksheet2.set_column('O:O', None, date_format)

#worksheet3.set_column('G:G', None, date_format)
#worksheet3.set_column('J:J', None, date_format)
#worksheet3.set_column('M:M', None, date_format)
#worksheet3.set_column('O:O', None, date_format)

#worksheet4.set_column('G:G', None, date_format)
#worksheet4.set_column('J:J', None, date_format)
#worksheet4.set_column('M:M', None, date_format)
#worksheet4.set_column('O:O', None, date_format)

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

for col_num, value in enumerate(dfPhaiTraTaiChuaDenHan.columns.values):
    worksheet4.write(0, col_num + 1, value, header_format)

border_fmt = workbook.add_format({'bottom':4, 'top':4, 'left':2, 'right':2})
worksheet1.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(dfPhaiTraTai3), len(dfPhaiTraTai3.columns)), {'type': 'no_errors', 'format': border_fmt})
worksheet2.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(dfPhaiTraTaiDuDieuKien), len(dfPhaiTraTaiDuDieuKien.columns)), {'type': 'no_errors', 'format': border_fmt})
worksheet3.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(dfPhaiTraTaiChuaDuDieuKien), len(dfPhaiTraTaiChuaDuDieuKien.columns)), {'type': 'no_errors', 'format': border_fmt})
worksheet4.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(dfPhaiTraTaiChuaDenHan), len(dfPhaiTraTaiChuaDenHan.columns)), {'type': 'no_errors', 'format': border_fmt})

worksheet1.freeze_panes(1, 0)
worksheet2.freeze_panes(1, 0)
worksheet3.freeze_panes(1, 0)
worksheet4.freeze_panes(1, 0)

writer.save()

# Send attachment via email
import keyring
import yagmail
from premailer import transform
yagmail.register('united.insurance.vn@gmail.com', 'Abc@123456')

yag = yagmail.SMTP('united.insurance.vn@gmail.com')
to_next = ['tuandq@uicvn.com']
subject_next = 'Công nợ tái bảo hiểm đến hạn '+str(dt.datetime.now().date())
body_next = f'Công nợ tái bảo hiểm đến hạn {str(dt.datetime.now().date())} <div>'
yag.send(to = to_next, subject = subject_next, contents = body_next, attachments=filename)
#End