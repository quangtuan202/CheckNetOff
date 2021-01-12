import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import filedialog as fd
from numpy.core.defchararray import startswith
from numpy.core.numeric import outer
import pyxlsb
import pandas as pd
import numpy as np
import subprocess
import numpy as np





win = tk.Tk()
win.title("Đối chiếu Paid list")
#-------------------------------------------------------------Declare global variables ------------------------------

paidListFile = ''
paidListFileUseCol=['Paid dated','Số hồ sơ BT/Số đơn','Currency','Số tiền']
paidListFileUseColDataType={'Paid dated': np.datetime64, 'Note': object , 'Currency' : object, 'Amount': np.float64 }

accFile = ''
accFileUseCol=['Ngày chứng từ','Số chứng từ','Số đơn', 'Số đơn HT', 'Số đơn ĐT','Loại tiền', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND']
accFileUseColDataType={'Ngày chứng từ':np.datetime64,'TK Nợ':object,'Loại tiền': object, 'TK Có':object, 'Số tiền':np.float64, 'Số tiền VND':np.float64}
savingFolder = ''
savingFileName=''
#radPremiumOrOthers=''
dfPaidList=pd.DataFrame
dfAcc=pd.DataFrame
dfAccFinal=pd.DataFrame
dfPaidListFinal=pd.DataFrame
dfPaidListGroupBy=pd.DataFrame
dfMerge=pd.DataFrame
dfAccGroupByVoucher=pd.DataFrame
dfAccFinalVoucher=pd.DataFrame

fullOrLeft=''

#------------------------------------------------------------- Widget event handle------------------------------------


def click_select_cpc():
    global paidListFile
    paidListFile = fd.askopenfilename()


def click_select_acc():
    global accFile
    accFile = fd.askopenfilename()


def click_save_file():
    global savingFolder
    savingFolder = fd.askdirectory()
    #print(saving_folder)

def button_open_saved_file_click():
    subprocess.Popen([f"{savingFolder}/{savingFileName.get()}.csv"], shell=True)
    win.destroy()

def rad_select_full_click():
    global fullOrLeft
    fullOrLeft = RadFullOrLeft.get() # return 'full'


def rad_select_left_click():
    global fullOrLeft
    fullOrLeft = RadFullOrLeft.get() # return 'left'

def return_policy(soDonHachToan, soDonDoiTru, soDon, taiKhoanNo, taiKhoanCo):
    if len(soDonHachToan) < 2 and len(soDonDoiTru) < 2:
        return soDon
    elif len(soDonHachToan) < 2 and len(soDon) < 2:
        return soDonDoiTru
    elif len(soDon) < 2 and len(soDonDoiTru) < 2:
        return soDonHachToan
    elif taiKhoanCo.startswith('112'): # TK no -> 'So don DT', TK co -> 'So don HT'
        if len(soDonDoiTru)>2:
            return soDonDoiTru
        else:
            return soDon
    elif taiKhoanNo.startswith('112'): # TK no -> 'So don DT', TK co -> 'So don HT'
        if len(soDonHachToan)>2:
            return soDonHachToan
        else:
            return soDon
    else:
        return


def return_account(taiKhoanNo,taiKhoanCo):
    if taiKhoanNo.startswith('112'):
        return taiKhoanNo
    elif taiKhoanCo.startswith('112'):
        return taiKhoanCo
    else:
        return 'TaiKhoanKhac'


def return_account_du(taiKhoanNo,taiKhoanCo):
    if taiKhoanNo.startswith('112'):
        return taiKhoanCo
    elif taiKhoanCo.startswith('112'):
        return taiKhoanNo
    else:
        return 'TaiKhoanKhac'

def return_amount(taiKHoanNo,taiKhoanCo,soTien):
    if taiKHoanNo.startswith('112'):
        return soTien
    elif taiKhoanCo.startswith('112'):
        return -soTien
    else:
        return 0



def click_execute():
    from tkinter import messagebox as msg
    global dfAccFinal
    global dfPaidListFinal
    global dfPaidList
    global dfAcc
    global dfPaidListGroupBy
    global dfAccGroupByVoucher
    global dfAccFinalVoucher
    global dfMerge
    global paidListFile
    global accFile
    global savingFolder
    global savingFileName

    #name=['No','Paid dated','Client Name','Amount','Currency','Type','Bank/Cash','Note','Số tiền','Place number','Số hồ sơ BT/Số đơn','Kỳ TT','Mã KH','Tài khoản']
    try:
        if paidListFile.endswith('xlsb'):
            dfPaidList = pd.read_excel(paidListFile, usecols=paidListFileUseCol,sheet_name='Claim paid', skiprows=1,engine='pyxlsb')
        else:
            dfPaidList = pd.read_excel(paidListFile, usecols=paidListFileUseCol,sheet_name='Claim paid', skiprows=1)
                  
        if accFile.endswith('xlsb'):
            dfAcc = pd.read_excel(accFile,usecols=accFileUseCol,skiprows=6,dtype=accFileUseColDataType, engine='pyxlsb')
        else:
            dfAcc = pd.read_excel(accFile,usecols=accFileUseCol,skiprows=6,dtype=accFileUseColDataType)
    except:
        msg.showerror("Error", "Only Excel files or CSV files are supported")
        pass

    values={'Số đơn':'0', 'Số đơn HT':'0','Số đơn ĐT':'0'}
    dfAcc=dfAcc.fillna(value=values)
    dfAcc=dfAcc.loc[dfAcc['TK Có'].str.startswith('112')]
    dfAcc['Policy']=dfAcc.apply(lambda x: return_policy(x['Số đơn'], x['Số đơn HT'], x['Số đơn ĐT'], x['TK Nợ'], x['TK Có']), axis=1)
    dfAcc['Tai khoan']=dfAcc.apply(lambda x: return_account(x['TK Nợ'], x['TK Có']), axis=1)  
    dfAcc['Tai khoan doi ung']=dfAcc.apply(lambda x: return_account_du(x['TK Nợ'], x['TK Có']), axis=1)  
    dfAcc['So tien NT']=dfAcc.apply(lambda x: return_amount(x['TK Nợ'], x['TK Có'],x['Số tiền']), axis=1) 
    dfAcc['So tien VND']=dfAcc.apply(lambda x: return_amount(x['TK Nợ'], x['TK Có'],x['Số tiền VND']), axis=1) 
    dfAccFinal=dfAcc.loc[(dfAcc['Tai khoan']!='TaiKhoanKhac') & (dfAcc['Tai khoan doi ung'].str.startswith(('624','331','333','131','133','138','338'))),['Ngày chứng từ','Policy','Loại tiền','So tien NT']] 
    dfAccFinalGroupby=dfAccFinal.groupby(['Ngày chứng từ','Policy','Loại tiền'],as_index=False)['So tien NT'].sum()
    dfAccFinalVoucher=dfAcc.loc[(dfAcc['Tai khoan']!='TaiKhoanKhac') & (dfAcc['Tai khoan doi ung'].str.startswith(('624','331','333','131','133','138','338'))),['Ngày chứng từ','Policy','Loại tiền','Số chứng từ']]
    dfAccFinalVoucher['So chung tu']=dfAccFinalVoucher.groupby(['Ngày chứng từ','Policy','Loại tiền'],as_index=False)['Số chứng từ'].transform(lambda x: ','.join(x))
    dfAccGroupByVoucher=dfAccFinalVoucher.drop_duplicates()
    dfPaidList=dfPaidList[dfPaidList['Số hồ sơ BT/Số đơn'].notna()]
    a=np.datetime64('2000-01-01')
    values={'Paid dated':a ,'Số hồ sơ BT/Số đơn': '0','Currency':'0','Số tiền':0}
    dfPaidList=dfPaidList.fillna(value=values)
    dfPaidListGroupBy=dfPaidList.groupby(['Paid dated','Số hồ sơ BT/Số đơn','Currency'],as_index=False)['Số tiền'].sum()
    if fullOrLeft=='full':
        dfMerge=dfPaidListGroupBy.merge(dfAccFinalGroupby,how='outer',left_on=['Paid dated','Số hồ sơ BT/Số đơn','Currency'],right_on=['Ngày chứng từ','Policy','Loại tiền'],suffixes=['PL','ACC']).merge(dfAccGroupByVoucher,how='outer',left_on=['Paid dated','Số hồ sơ BT/Số đơn','Currency'],right_on=['Ngày chứng từ','Policy','Loại tiền'],suffixes=['PL2','ACC2'])
    else:
        dfMerge=dfPaidListGroupBy.merge(dfAccFinalGroupby,how='left',left_on=['Paid dated','Số hồ sơ BT/Số đơn','Currency'],right_on=['Ngày chứng từ','Policy','Loại tiền'],suffixes=['PL','ACC']).merge(dfAccGroupByVoucher,how='left',left_on=['Paid dated','Số hồ sơ BT/Số đơn','Currency'],right_on=['Ngày chứng từ','Policy','Loại tiền'],suffixes=['PL2','ACC2'])
    dfMerge['Chenh lech']=dfMerge['Số tiền']+dfMerge['So tien NT']
    dfMerge=dfMerge.loc[:,['Paid dated','Số hồ sơ BT/Số đơn','Currency','Số tiền','Ngày chứng từPL2','PolicyPL2','Loại tiềnPL2','So tien NT','Chenh lech','So chung tu']]
    dfMerge.to_csv(f"{savingFolder}/{savingFileName.get()}.csv", encoding="utf-8-sig", index=False)
    msg.showinfo("Information", "Đã hoàn thành")

    subprocess.Popen([f"{savingFolder}/{savingFileName.get()}.csv"], shell=True)

    
    paidListFile = ''
    accFile = ''
    savingFolder = ''
    savingFileName=''
    #radPremiumOrOthers=''
    #dfPaidList=pd.DataFrame
    #dfAcc=pd.DataFrame
    #dfAccFinal=pd.DataFrame
    #dfPaidListFinal=pd.DataFrame
    #dfPaidListGroupBy=pd.DataFrame


# ----------------------------------------------------------------- Widget------------------------------------------------

RadFullOrLeft = tk.StringVar()


rad_select_full = ttk.Radiobutton(win, text='Full', width=30, value='full', variable=RadFullOrLeft, command=rad_select_full_click)
rad_select_full.grid(column=0, row=0)

rad_select_claim = ttk.Radiobutton(win, text='Left', width=30, value='left', variable=RadFullOrLeft, command=rad_select_left_click)
rad_select_claim.grid(column=1, row=0)

ttk.Label(win, text="Chọn file Paid list", width=20).grid(column=0, row=1)
button1 = ttk.Button(win, text="Paid list", width=20, command=click_select_cpc)
button1.grid(column=1, row=1)

ttk.Label(win, text="Chọn file CTGS", width=20).grid(column=0, row=2)
button2 = ttk.Button(win, text="CTGS", width=20, command=click_select_acc)
button2.grid(column=1, row=2)

ttk.Label(win, text="Chọn thư mục lưu file", width=20).grid(column=0, row=3)
button2 = ttk.Button(win, text="Folder", width=20, command=click_save_file)
button2.grid(column=1, row=3)

ttk.Label(win, text="Nhập tên file", width=20).grid(column=0, row=4)
savingFileName = ttk.Entry(win, width=20)
savingFileName.grid(column=1, row=4)

ttk.Label(win, text="Xử lý số liệu", width=20).grid(column=0, row=5)
button3 = ttk.Button(win, text="Run", width=20, command=click_execute)
button3.grid(column=1, row=5)

label_open_saved_file = ttk.Label(win, text="Mở file", width=20).grid(column=0, row=6)
button_open_saved_file = ttk.Button(win, text="Mở file", width=20, command=button_open_saved_file_click)
button_open_saved_file.grid(column=1, row=6)




win.mainloop()