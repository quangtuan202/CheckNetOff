import pyxlsb
import pandas as pd
import numpy as np
import os
import subprocess

accCol=['Stt','NhomChungTu','LoaiChungTu','SoChungTu',
        'NgayChungTu','NgayHieuLuc','ChiNhanh','Phong',
        'SoDon','SoRef','SoDonHT','SoDonDT','ACRef','MaDTGiaoDich',
        'TenDTGiaoDich','MaKhachHang','TenKhachHang','MaKhachHangDoiTru',
        'TenKhachHangDoiTru','NoiDung','TaiKhoanNo','TaiKhoanCo','SoTien',
        'LoaiTien','SoTienVND','TyGia','MaNghiepVu','SoSeri','NgayHoaDon',
        'MaSoThue','SoHoaDon','NoiDung','GhiChu','NguoiHachToan','NguoiXuatHoaDon']
##################################################################################################
import pyxlsb
import pandas as pd
import numpy as np
import subprocess

##################################################################################################
class GLdataframe:
    def __init__(self,folder,useCol,account,account_type):
        #self.useCol=useCol
        self.account=account
        self.account_type=account_type
        df_list=[]
        for fileName in os.listdir(folder):
            if fileName.endswith('xlsb'):
                self.dataframe=pd.read_excel(f'{folder}/{fileName}',sheet_name=None,header=None, engine='pyxlsb')
            else:
                self.dataframe=pd.read_excel(f'{folder}/{fileName}',sheet_name=None,header=None,)

            dfList=list(self.dataframe.values())
            dfList[0]=dfList[0].iloc[7:,:]
            df_list.extend(dfList)
        df=pd.concat(df_list, axis=0)
        df=df.iloc[:,1:]
        df.columns=accCol
        self.dataframe=df.loc[df['TaiKhoanNo'].isin(account)|df['TaiKhoanCo'].isin(account),useCol]
        self.dataframe['policy']=self.dataframe.apply(lambda x: self.return_policy(x['SoDon'],x['SoDonHT'],x['SoDonDT'],x['TaiKhoanNo'],x['TaiKhoanCo']),axis=1)
        self.dataframe[['SoChungTu','SoDon','SoDonHT','SoDonDT','TaiKhoanNo','TaiKhoanCo']]=self.dataframe[['SoChungTu','SoDon','SoDonHT','SoDonDT','TaiKhoanNo','TaiKhoanCo']].fillna('').astype(str)
    ####################################################################        
    def return_policy(self,don,don_ht, don_dt,debit_account, credit_account):
        if len(don_ht) < 2 and len(don_dt) < 2:
            return don
        elif len(don_ht) < 2 and len(don) < 2:
            return don_dt
        elif len(don_dt) < 2 and len(don) < 2:
            return don_ht
        elif debit_account in self.account:
            return don_dt
        elif credit_account in self.account:
            return don_ht
        else:
            return
    ####################################################################  
    def return_amount(self,debit_account, credit_account, amount):
        if self.account_type=='debit':
            if credit_account in self.acc_combined_list:
                return -amount
            elif debit_account in self.acc_combined_list:
                return amount
            else:
                return
        else: # account_type='credit'
            if credit_account in self.acc_combined_list:
                return amount
            elif debit_account in self.acc_combined_list:
                return -amount
            else:
                return

##################################################################################################
class account:
    def __init__(self,fileName,useCol,vatType,branch,phanLoaiDoanhThu,UIC_Dong):
        #vatType: 0,10,free,all
        #branch: HANOI, HOCHIMINH, DANANG, VINH
        #phanLoaiDoanhThu: PhiGoc, GiamPhiGoc, HoanPhiGoc, PhiCapDon, PhiGiamDinh, ThanhLyHang, DoanhThuKhac
        #UIC_Dong: UIC, Dong
        if fileName.endswith('csv'):
            self.dataframe=pd.read_csv(fileName, usecols=useCol)
        elif fileName.endswith('xlsb'):
            self.dataframe = pd.read_excel(fileName, usecols=useCol, engine='pyxlsb')
        elif fileName.endswith('xls') or fileName.endswith('xlsx'):
            self.dataframe = pd.read_excel(fileName, usecols=useCol)
        else: 
            msg.showinfo("Information", "Select Excel or CSV file only")
            pass
        condition = (self.dataframe['VAT_Type']==vatType)&(self.dataframe['Branch']==branch)&(self.dataframe['PhanLoaiDoanhThu']==phanLoaiDoanhThu)
        self.dataframe=self.dataframe.loc[condition]
        self.dataframe=self.dataframe.loc[:,['Account']]
        self.series=self.dataframe['Account'].values

##################################################################################################
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

