# 07/04/2020

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import filedialog as fd
import pyxlsb
import pandas as pd
import numpy as np
import subprocess

win = tk.Tk()
win.title("UIC ACC-CPC Reconciliation")

# ------------------------------------------------------------------------------------
cpc_file = ''
acc_file = ''
saving_folder = ''
#acc_list=[]
use_cols_cpc=[]
#use_cols_acc=[]


acc_list = [511111, 511112, 511113, 511114, 511115, 511116, 511131, 511133, 511134, 511136, 511211,
                    531111, 531112, 531113, 531114, 531115, 531116, 531131, 531133, 531134, 531136, 531211,
                    532111, 532112, 532113, 532114, 532115, 532116, 532117]

premium_acc_list = [511111, 511112, 511113, 511114, 511115, 511116, 511131, 511133, 511134, 511136, 511211,
                    531111, 531112, 531113, 531114, 531115, 531116, 531131, 531133, 531134, 531136, 531211,
                    532111, 532112, 532113, 532114, 532115, 532116, 532117]


commission_acc_list = [624141,624143,624241,624173]
brokerage_acc_list= [624142,624144]
use_cols_premium_cpc=['Loại hình bảo hiểm', 'Số đơn/ Số Endor', 'Phí Bảo Hiểm\n(nguyên tệ)','Phí bảo hiểm\n(VND)']

use_cols_commission_cpc=['Số đơn/ Số Endor','Môi giới phí BHG / Hoa hồng NTBH','Hoa hồng đại lý bảo hiểm']
use_cols_claim_cpc=['Số hồ sơ bồi thường\t','Bồi thường\n(VND)','Phí giám định']
use_cols_tpa_cpc=['Số đơn/ Số Endor','Số tiền TPA']
use_cols_reinsurance_cpc=['Số đơn/ Số Endor','Phí nhượng TBH FAC']

use_cols_acc=['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND']

# 624141 : Hoa hồng đại lý
# 624143 : Hoa hong Dai ly (tra cty Leader)
# 624173 : Chi khen thuong Dai ly
# 624241 : Hoa hồng nhận tái bảo hiểm 

# 624142 : Môi giới phí
# 624144 : Moi gioi phi (tra cty Leader)


claim_acc_list = [513811,513812,624111,624112,624118,624162,624211]





# 624112 : Chi giam dinh boi thuong
# 513811 : HN- Doanh thu hoạt động thanh lý hàng tổn thất
# 513812 :HCMC- Doanh thu hoạt động thanh lý hàng tổn thất



reinsurance_acc_list = [533111,511411 ]

tpa_acc_list = [624181]


df_cpc = pd.DataFrame()
df_acc = pd.DataFrame()
df_unique_policies=pd.DataFrame()

# ------------------------------------------------------------------------------------

def click_select_cpc():
    global cpc_file
    cpc_file = fd.askopenfilename()


def click_select_acc():
    global acc_file
    acc_file = fd.askopenfilename()


def click_save_file():
    global saving_folder
    saving_folder = fd.askdirectory()
    #print(saving_folder)

def button_open_saved_file_click():
    subprocess.Popen([f"{saving_folder}/{saving_file_name.get()}.csv"], shell=True)
    win.destroy()






def click_execute():
    from tkinter import messagebox as msg
    global df_cpc
    global df_acc

    if cpc_file.endswith('csv'):
        df_cpc = pd.read_csv(cpc_file, usecols=['Loại hình bảo hiểm', 'Số đơn/ Số Endor', 'Phí Bảo Hiểm\n(nguyên tệ)',
                                         'Phí bảo hiểm\n(VND)'])
    elif cpc_file.endswith('xlsb'):
        df_cpc = pd.read_excel(cpc_file, usecols=['Loại hình bảo hiểm', 'Số đơn/ Số Endor', 'Phí Bảo Hiểm\n(nguyên tệ)',
                                           'Phí bảo hiểm\n(VND)'], engine='pyxlsb')
    else:
        df_cpc = pd.read_excel(cpc_file, usecols=['Loại hình bảo hiểm', 'Số đơn/ Số Endor', 'Phí Bảo Hiểm\n(nguyên tệ)',
                                           'Phí bảo hiểm\n(VND)'])

    if acc_file.endswith('csv'):
        df_acc = pd.read_csv(acc_file,
                             usecols=['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND'],
                             encoding="utf-8").fillna('0')
        # df_acc['Số tiền']=df_acc['Số tiền'].astype('np.float32')
        # df_acc['Số tiền']=df_acc['Số tiền VND'].astype('float64')
    elif acc_file.endswith('xlsb'):
        df_acc = pd.read_excel(acc_file,
                               usecols=['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND'],
                               encoding="utf-8", engine='pyxlsb').fillna('0')
        # df_acc['Số tiền']=df_acc['Số tiền'].astype('float64')
        # df_acc['Số tiền']=df_acc['Số tiền VND'].astype('float64')
    else:
        df_acc = pd.read_excel(acc_file,
                               usecols=['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND'],
                               encoding="utf-8").fillna('0')
        # df_acc['Số tiền']=df_acc['Số tiền'].astype('float64')
        # df_acc['Số tiền']=df_acc['Số tiền VND'].astype('float64')

    # --------------------------------------------------------------------------------
    df_cpc.rename(columns={'Số đơn/ Số Endor': 'Policy', 'Phí Bảo Hiểm\n(nguyên tệ)': 'NT_cpc',
                           'Phí bảo hiểm\n(VND)': 'VND_cpc'}, inplace=True)
    # apply function 'dept', axis = 1 for all rows
    #values = {'NT_cpc': 0, 'NT_acc': 0, 'VND_cpc': 0, 'VND_acc': 0}
    #df_cpc.fillna(value=values, inplace=True)

    df_cpc = df_cpc.groupby('Policy', as_index=False).sum()
    #df_cpc['Dept'] = df_cpc.apply(lambda x: dept(x['Policy_cpc']), axis=1)
    #df_cpc['Lob'] = df_cpc.apply(lambda x: policy_type(x['Policy_cpc']), axis=1)
    df_cpc['Policy_dot'] = df_cpc.apply(lambda x: policy_dot(x['Policy']), axis=1)
    df_cpc['Policy_hyphen'] = df_cpc.apply(lambda x: policy_hyphen(x['Policy']), axis=1)
    df_cpc['Policy_en'] = df_cpc.apply(lambda x: policy_en(x['Policy']), axis=1)
    df_cpc['Policy_dot_hyphen_en'] = df_cpc.apply(
        lambda x: policy_dot_hyphen_en(x['Policy_dot'], x['Policy_hyphen'], x['Policy_en']), axis=1)
    df_cpc['Policy_dept']=df_cpc['Policy'].apply(policy_dept)

    df_cpc_dot = df_cpc[['Policy_dot', 'NT_cpc', 'VND_cpc']].copy()
    df_cpc_dot = df_cpc_dot.groupby('Policy_dot', as_index=False).sum()  # as_index=False to show grouped column
    df_cpc_dot.rename(columns={'NT_cpc': 'NT_cpc_dot', 'VND_cpc': 'VND_cpc_dot'}, inplace=True)

    df_cpc_hyphen = df_cpc[['Policy_hyphen', 'NT_cpc', 'VND_cpc']].copy()
    df_cpc_hyphen = df_cpc_hyphen.groupby('Policy_hyphen', as_index=False).sum()
    df_cpc_hyphen.rename(columns={'NT_cpc': 'NT_cpc_hyphen', 'VND_cpc': 'VND_cpc_hyphen'}, inplace=True)

    df_cpc_en = df_cpc[['Policy_en', 'NT_cpc', 'VND_cpc']].copy()
    df_cpc_en = df_cpc_en.groupby('Policy_en', as_index=False).sum()
    df_cpc_en.rename(columns={'NT_cpc': 'NT_cpc_en', 'VND_cpc': 'VND_cpc_en'}, inplace=True)

    df_cpc_dot_hyphen_en = df_cpc[['Policy_dot_hyphen_en', 'NT_cpc', 'VND_cpc']].copy()
    df_cpc_dot_hyphen_en = df_cpc_dot_hyphen_en.groupby('Policy_dot_hyphen_en', as_index=False).sum()
    df_cpc_dot_hyphen_en.rename(columns={'NT_cpc': 'NT_cpc_dot_hyphen_en', 'VND_cpc': 'VND_cpc_dot_hyphen_en'}, inplace=True)

    df_cpc_dept = df_cpc[['Policy_dept', 'NT_cpc', 'VND_cpc']].copy()
    df_cpc_dept = df_cpc_dept.groupby('Policy_dept', as_index=False).sum()
    df_cpc_dept.rename(columns={'NT_cpc': 'NT_cpc_dept', 'VND_cpc': 'VND_cpc_dept'}, inplace=True)
    # --------------------------------------------------------------------------------

    df_acc['Policy'] = df_acc.apply(
        lambda x: return_policy(x['Số đơn'], x['Số đơn HT'], x['Số đơn ĐT'], x['TK Nợ'], x['TK Có']), axis=1)
    df_acc['NT_acc'] = df_acc.apply(lambda x: return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền']), axis=1)
    df_acc['VND_acc'] = df_acc.apply(lambda x: return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền VND']), axis=1)
    df_acc = df_acc[['Policy','NT_acc','VND_acc']]
    #df_acc = df_acc.drop(['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND'], axis=1)
    df_acc = df_acc[~pd.isnull(df_acc['VND_acc'])]

    #values = {'NT_cpc': 0, 'NT_acc': 0, 'VND_cpc': 0, 'VND_acc': 0}
    #df_acc.fillna(value=values, inplace=True)

    df_acc = df_acc.groupby('Policy', as_index=False).sum()
    df_acc['Policy_dot'] = df_acc.apply(lambda x: policy_dot(x['Policy']), axis=1)
    df_acc['Policy_hyphen'] = df_acc.apply(lambda x: policy_hyphen(x['Policy']), axis=1)
    df_acc['Policy_en'] = df_acc.apply(lambda x: policy_en(x['Policy']), axis=1)
    df_acc['Policy_dot_hyphen_en'] = df_acc.apply(
        lambda x: policy_dot_hyphen_en(x['Policy_dot'], x['Policy_hyphen'], x['Policy_en']), axis=1)
    df_acc['Policy_dept']=df_acc['Policy'].apply(policy_dept)

    df_acc_dot = df_acc[['Policy_dot', 'NT_acc', 'VND_acc']].copy()
    df_acc_dot = df_acc_dot.groupby('Policy_dot', as_index=False).sum()  # as_index=False to show grouped column
    df_acc_dot.rename(columns={'NT_acc': 'NT_acc_dot', 'VND_acc': 'VND_acc_dot'}, inplace=True)

    df_acc_hyphen = df_acc[['Policy_hyphen', 'NT_acc', 'VND_acc']].copy()
    df_acc_hyphen = df_acc_hyphen.groupby('Policy_hyphen', as_index=False).sum()
    df_acc_hyphen.rename(columns={'NT_acc': 'NT_acc_hyphen', 'VND_acc': 'VND_acc_hyphen'}, inplace=True)

    df_acc_en = df_acc[['Policy_en', 'NT_acc', 'VND_acc']].copy()
    df_acc_en = df_acc_en.groupby('Policy_en', as_index=False).sum()
    df_acc_en.rename(columns={'NT_acc': 'NT_acc_en', 'VND_acc': 'VND_acc_en'}, inplace=True)

    df_acc_dot_hyphen_en = df_acc[['Policy_dot_hyphen_en', 'NT_acc', 'VND_acc']].copy()
    df_acc_dot_hyphen_en = df_acc_dot_hyphen_en.groupby('Policy_dot_hyphen_en', as_index=False).sum()
    df_acc_dot_hyphen_en.rename(columns={'NT_acc': 'NT_acc_dot_hyphen_en', 'VND_acc': 'VND_acc_dot_hyphen_en'}, inplace=True)

    df_acc_dept = df_acc[['Policy_dept', 'NT_acc', 'VND_acc']].copy()
    df_acc_dept = df_acc_dept.groupby('Policy_dept', as_index=False).sum()
    df_acc_dept.rename(columns={'NT_acc': 'NT_acc_dept', 'VND_acc': 'VND_acc_dept'}, inplace=True)
   

    #-------------------------------------------------------------------------------------
    # Create a list of unique items
    df_unique_policies=unique_item(df_cpc['Policy'],df_acc['Policy'])
    df_unique_policies['Policy_dot']=df_unique_policies['Policy'].apply(policy_dot)
    df_unique_policies['Policy_hyphen']=df_unique_policies['Policy'].apply(policy_hyphen)
    df_unique_policies['Policy_en']=df_unique_policies['Policy'].apply(policy_en)
    df_unique_policies['Policy_dot_hyphen_en']=df_unique_policies.apply(
        lambda x: policy_dot_hyphen_en(x['Policy_dot'], x['Policy_hyphen'], x['Policy_en']), axis=1)
    df_unique_policies['Policy_dept']=df_unique_policies['Policy'].apply(policy_dept)
 #-------------------------------------------------------------------------------------
    # Join unique list with other dataframes
    df_merge_all=df_unique_policies.merge(df_cpc,how='left',on='Policy').merge(df_acc,how='left',on='Policy').merge(df_cpc_dot,how='left'
    ,on='Policy_dot').merge(df_acc_dot,how='left',on='Policy_dot').merge(df_cpc_hyphen,how='left',on='Policy_hyphen').merge(df_acc_hyphen,how='left'
    ,on='Policy_hyphen').merge(df_cpc_en,how='left',on='Policy_en').merge(df_acc_en,how='left',on='Policy_en').merge(df_cpc_dot_hyphen_en,how='left'
    ,on='Policy_dot_hyphen_en').merge(df_acc_dot_hyphen_en,how='left',on='Policy_dot_hyphen_en').merge(df_cpc_dept,how='left',on='Policy_dept').merge(df_acc_dept,how='left',on='Policy_dept')
    
    values={'NT_cpc':0,'NT_cpc_dot':0,'NT_cpc_hyphen':0,'NT_cpc_en':0,'NT_cpc_dot_hyphen_en':0,'NT_cpc_dept':0,'VND_cpc':0,'VND_cpc_dot':0,
           'VND_cpc_hyphen':0,'VND_cpc_en':0,'VND_cpc_dot_hyphen_en':0,'VND_cpc_dept':0,
           'NT_acc':0,'NT_acc_dot':0,'NT_acc_hyphen':0,'NT_acc_en':0,'NT_acc_dot_hyphen_en':0,
           'NT_acc_dept':0,'VND_acc':0,'VND_acc_dot':0,'VND_acc_hyphen':0,'VND_acc_en':0,'VND_acc_dot_hyphen_en':0,'VND_acc_dept':0}
    
    df_merge_all.fillna(value=values,inplace=True)
    df_merge_all['Chenh lech NT']=df_merge_all['NT_cpc']-df_merge_all['NT_acc']
    df_merge_all['Chenh lech NT_dot']=df_merge_all['NT_cpc_dot']-df_merge_all['NT_acc_dot']
    df_merge_all['Chenh lech NT_hyphen']=df_merge_all['NT_cpc_hyphen']-df_merge_all['NT_acc_hyphen']
    df_merge_all['Chenh lech NT_en']=df_merge_all['NT_cpc_en']-df_merge_all['NT_acc_en']
    df_merge_all['Chenh lech NT_dot_hyphen_en']=df_merge_all['NT_cpc_dot_hyphen_en']-df_merge_all['NT_acc_dot_hyphen_en']
    df_merge_all['Chenh lech NT_dept']=df_merge_all['NT_cpc_dept']-df_merge_all['NT_acc_dept']

    df_merge_all['Chenh lech VND']=df_merge_all['VND_cpc']-df_merge_all['VND_acc']
    df_merge_all['Chenh lech VND_dot']=df_merge_all['VND_cpc_dot']-df_merge_all['VND_acc_dot']
    df_merge_all['Chenh lech VND_hyphen']=df_merge_all['VND_cpc_hyphen']-df_merge_all['VND_acc_hyphen']
    df_merge_all['Chenh lech VND_en']=df_merge_all['VND_cpc_en']-df_merge_all['VND_acc_en']
    df_merge_all['Chenh lech VND_dot_hyphen_en']=df_merge_all['VND_cpc_dot_hyphen_en']-df_merge_all['VND_acc_dot_hyphen_en']
    df_merge_all['Chenh lech VND_dept']=df_merge_all['VND_cpc_dept']-df_merge_all['VND_acc_dept']

    df_merge_all['Chenh lech NT_abs']= df_merge_all['Chenh lech NT'].abs()
    df_merge_all['Chenh lech NT_dot_abs']= df_merge_all['Chenh lech NT_dot'].abs()
    df_merge_all['Chenh lech NT_hyphen_abs']= df_merge_all['Chenh lech NT_hyphen'].abs()
    df_merge_all['Chenh lech NT_en_abs']= df_merge_all['Chenh lech NT_en'].abs()
    df_merge_all['Chenh lech NT_dot_hyphen_en_abs']= df_merge_all['Chenh lech NT_dot_hyphen_en'].abs()
    df_merge_all['Chenh lech NT_dept_abs']= df_merge_all['Chenh lech NT_dept'].abs()

    df_merge_all['Chenh lech VND_abs']= df_merge_all['Chenh lech VND'].abs()
    df_merge_all['Chenh lech VND_dot_abs']= df_merge_all['Chenh lech VND_dot'].abs()
    df_merge_all['Chenh lech VND_hyphen_abs']= df_merge_all['Chenh lech VND_hyphen'].abs()
    df_merge_all['Chenh lech VND_en_abs']= df_merge_all['Chenh lech VND_en'].abs()
    df_merge_all['Chenh lech VND_dot_hyphen_en_abs']= df_merge_all['Chenh lech VND_dot_hyphen_en'].abs()
    df_merge_all['Chenh lech VND_dept_abs']= df_merge_all['Chenh lech VND_dept'].abs()
   
    df_merge_all['Min_Diff_NT'] = df_merge_all.apply(
        lambda x: return_min(x['Chenh lech NT_abs'], x['Chenh lech NT_dot_abs'], x['Chenh lech NT_hyphen_abs'], x['Chenh lech NT_en_abs'],
                             x['Chenh lech NT_dot_hyphen_en_abs'], x['Chenh lech NT_dept_abs'], x['Chenh lech NT'], x['Chenh lech NT_dot'], x['Chenh lech NT_hyphen'], x['Chenh lech NT_en'],
                             x['Chenh lech NT_dot_hyphen_en'], x['Chenh lech NT_dept']), axis=1)
    df_merge_all['Min_Diff_VND'] = df_merge_all.apply(
        lambda x: return_min(x['Chenh lech VND_abs'], x['Chenh lech VND_dot_abs'], x['Chenh lech VND_hyphen_abs'], x['Chenh lech VND_en_abs'],
                             x['Chenh lech VND_dot_hyphen_en_abs'], x['Chenh lech VND_dept_abs'], x['Chenh lech VND'], x['Chenh lech VND_dot'], x['Chenh lech VND_hyphen'], x['Chenh lech VND_en'],
                             x['Chenh lech VND_dot_hyphen_en'], x['Chenh lech VND_dept']), axis=1)

    #-------------------------------------------------------------------------------------


    df_merge_all.to_csv(f"{saving_folder}/{saving_file_name.get()}.csv", encoding="utf-8-sig", index=False)
    # df_merge.to_csv(c+r'/result5.csv',encoding="utf-8-sig",index=False)
    msg.showinfo("Information", "Đã hoàn thành")
    win.destroy()  # Close window after execution


# --------------------------------------------------------------------------------

def read_file():
    try:
        from tkinter import messagebox as msg
        global df_cpc
        global df_acc
        global use_cols_cpc
        global use_cols_acc
        # "Revenue", "Commission", "Claim", 'TPA', 'Reinsurance Premium'

        if option_control_variable.get()=='Revenue':
            use_cols_cpc=use_cols_premium_cpc
        elif option_control_variable.get()=='Commission':
            use_cols_cpc=use_cols_commission_cpc
        elif option_control_variable.get()=='Claim':
            use_cols_cpc=use_cols_claim_cpc
        elif option_control_variable.get()=='TPA':
            use_cols_cpc=use_cols_tpa_cpc
        elif option_control_variable.get()=='Reinsurance Premium':
            use_cols_cpc=use_cols_reinsurance_cpc
        


            if cpc_file.endswith('csv'):
                df_cpc = pd.read_csv(cpc_file, usecols=['Loại hình bảo hiểm', 'Số đơn/ Số Endor', 'Phí Bảo Hiểm\n(nguyên tệ)',
                                         'Phí bảo hiểm\n(VND)'])
            elif cpc_file.endswith('xlsb'):
                df_cpc = pd.read_excel(cpc_file, usecols=['Loại hình bảo hiểm', 'Số đơn/ Số Endor', 'Phí Bảo Hiểm\n(nguyên tệ)',
                                           'Phí bảo hiểm\n(VND)'], engine='pyxlsb')
            else:
                df_cpc = pd.read_excel(cpc_file, usecols=['Loại hình bảo hiểm', 'Số đơn/ Số Endor', 'Phí Bảo Hiểm\n(nguyên tệ)',
                                           'Phí bảo hiểm\n(VND)'])

            if acc_file.endswith('csv'):
                df_acc = pd.read_csv(acc_file,
                             usecols=['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND'],
                             encoding="utf-8").fillna('0')
        # df_acc['Số tiền']=df_acc['Số tiền'].astype('np.float32')
        # df_acc['Số tiền']=df_acc['Số tiền VND'].astype('float64')
            elif acc_file.endswith('xlsb'):
                df_acc = pd.read_excel(acc_file,
                               usecols=['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND'],
                               encoding="utf-8", engine='pyxlsb').fillna('0')
        # df_acc['Số tiền']=df_acc['Số tiền'].astype('float64')
        # df_acc['Số tiền']=df_acc['Số tiền VND'].astype('float64')
            else:
                df_acc = pd.read_excel(acc_file,
                               usecols=['Số đơn', 'Số đơn HT', 'Số đơn ĐT', 'TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND'],
                               encoding="utf-8").fillna('0')
        # df_acc['Số tiền']=df_acc['Số tiền'].astype('float64')
        # df_acc['Số tiền']=df_acc['Số tiền VND'].astype('float64')
        
        return df_cpc
        return df_acc
        
    except:
        msg.showerror("Error", "Only Excel files or CSV files are supported")
        pass


# --------------------------------------------------------------------------------
# Make a DataFrame of unique items
def unique_item(a=pd.DataFrame(),b=pd.DataFrame()):
    df=pd.concat([a,b]).unique()
    df_unique=pd.DataFrame(df,columns=['Policy'])
    return df_unique

# --------------------------------------------------------------------------------

def dept(codename):
    if 'HY' in codename:
        return 'NKSJ - HN'
    elif not '.HN' in codename and 'HN' in codename:
        return 'NKSJ - HN'
    elif 'HL' in codename:
        return 'NKSJ - HN'
    elif 'SY' in codename:
        return 'NKSJ - HCMC'
    elif 'SN' in codename:
        return 'NKSJ - HCMC'
    elif 'HB' in codename:
        return 'BM - HN'
    elif 'SB' in codename:
        return 'BM - HCMC'
    elif 'HP' in codename:
        return 'Retail (HN)'
    elif 'HR' in codename:
        return 'Retail (HN)'
    elif 'SR' in codename:
        return 'Retail (HCMC)'
    elif 'HG' in codename:
        return 'KB - HN'
    elif 'HUR' in codename:
        return 'Reins. Dept.'
    elif 'HUD' in codename:
        return 'Reins. Dept.'
    elif 'HBR' in codename:
        return 'BM - HN'
    elif 'DR' in codename:
        return 'Danang BR'
    elif 'VR' in codename:
        return 'Vinh BR'
    elif 'SG' in codename:
        return 'KB - HCMC'
    elif codename[10:12] == '11':
        return 'Marketing 1 (BM) - HN '
    elif codename[10:12] == '12':
        return 'Marketing 1 (BM) - HCM'
    elif codename[10:12] == '21':
        return 'Marketing 2 (SPJ) - HN'
    elif codename[10:12] == '22':
        return 'Marketing 2 (SPJ) - HCM'
    elif codename[10:12] == '32':
        return 'Phòng cấp đơn HCM'
    elif codename[10:12] == '41':
        return 'Retail (HN)'
    elif codename[10:12] == '42':
        return 'Retail (HCMC)'
    elif codename[10:12] == '01':
        return 'Reins. Dept.'
    elif codename[10:12] == '03':
        return 'Retail (HN)'
    elif codename[10:12] == '13':
        return 'Danang BR'
    elif codename[10:12] == '14':
        return 'Vinh BR'
    elif codename[10:12] == '31':
        return 'Retail (HN)'
    elif codename[10:12] == '51':
        return 'Retail (HN)'
    elif codename[10:12] == '16':
        return 'Retail (HN)'
    elif codename[10:12] == '52':
        return 'Retail (HCMC)'
    else:
        return 'undefined'


# --------------------------------------------------------------------------------

def policy_type(policy):
    if '501' in policy and policy.startswith('00'):
        return 'MOT'
    elif '502' in policy and policy.startswith('00'):
        return 'CTP'
    elif 'CTP' in policy and policy.startswith('0'):
        return 'CTP'
    elif 'MOT' in policy and policy.startswith('0'):
        return 'MOT'
    elif 'MOX' in policy and policy.startswith('0'):
        return 'MOX'
    elif 'ARRTRIP' in policy:
        return policy[8:11]
    else:
        return policy[0:3]


# --------------------------------------------------------------------------------

def policy_dot(policy):
    if policy.startswith('00') and policy.endswith('T') and policy.find(".") == -1:
        return policy[:-1]
    elif policy.startswith('ARRTRIP.'):
        a=len(policy)+1
        return policy[8:a]
    elif policy.find(".") != -1:
        return policy[0:policy.find(".")]
    else:
        return policy


# --------------------------------------------------------------------------------

def policy_hyphen(policy):
    if policy.find("-") != -1:
        return policy[0:policy.find("-")]
    else:
        return policy


# --------------------------------------------------------------------------------

def policy_en(policy):
    if policy.find("EN") != -1:
        return policy[0:policy.find("EN")]
    else:
        return policy


# --------------------------------------------------------------------------------

def policy_dot_hyphen_en(a, b, c):
    if len(a) == min(len(a), len(b), len(c)):
        return a
    elif len(b) == min(len(a), len(b), len(c)):
        return b
    else:
        return c


# ---------------------------------------------------------------------------------
# if find() return -1, it found no value
def policy_dept(policy):
    if policy.find('HN') != -1:
        return policy[0:(policy.find('HN') + 3)]
    elif policy.find('HY') != -1:
        return policy[0:(policy.find('HY') + 3)]
    elif policy.find('SYCAR') != -1:
        return policy[0:(policy.find('SYCAR') + 5)]
    elif policy.find('SY') != -1:
        return policy[0:(policy.find('SY') + 3)]
    elif policy.find('SN') != -1:
        return policy[0:(policy.find('SN') + 3)]
    elif policy.find('HB') != -1:
        return policy[0:(policy.find('HB') + 3)]
    elif policy.find('SB') != -1:
        return policy[0:(policy.find('SB') + 3)]
    elif policy.find('HP') != -1:
        return policy[0:(policy.find('HP') + 3)]
    elif policy.find('HR') != -1:
        return policy[0:(policy.find('HR') + 3)]
    elif policy.find('SR') != -1:
        return policy[0:(policy.find('SR') + 3)]
    elif policy.find('HG') != -1:
        return policy[0:(policy.find('HG') + 3)]
    elif policy.find('HU') != -1:
        return policy[0:(policy.find('HU') + 3)]
    elif policy.find('DR') != -1:
        return policy[0:(policy.find('DR') + 3)]
    elif policy.find('VR') != -1:
        return policy[0:(policy.find('VR') + 3)]
    elif policy.find('SG') != -1:
        return policy[0:(policy.find('SG') + 3)]
    else:
        return policy




# --------------------------------------------------------------------------------

def return_amount(a, b, c):
    if a in acc_list:
        return -c
    elif b in acc_list:
        return c
    else:
        return


# --------------------------------------------------------------------------------

def return_policy(a, b, c, d, e):
    if len(a) < 2 and len(b) < 2:
        return c
    elif len(a) < 2 and len(c) < 2:
        return b
    elif len(b) < 2 and len(c) < 2:
        return a
    elif d in acc_list:
        if len(a) < 2:
            return c
        else:
            return a
    elif e in acc_list:
        if len(a) < 2:
            return b
        else:
            return a
    else:
        return


# --------------------------------------------------------------------------------
# a, b, c, d, e, f are abs value of colums
def return_min(a, b, c, d, e, f, a1, b1, c1, d1, e1, f1):
    if a == min(a, b, c, d, e, f):
        return a1
    elif b == min(a, b, c, d, e, f):
        return b1
    elif c == min(a, b, c, d, e, f):
        return c1
    elif d == min(a, b, c, d, e, f):
        return d1
    elif e == min(a, b, c, d, e, f):
        return e1
    elif f == min(a, b, c, d, e, f):
        return f1
    else:
        return min(a, b, c, d, e, f)


# --------------------------------------------------------------------------------
option_control_variable = tk.StringVar()
options = ['Select',"Revenue", "Commission", "Claim", 'TPA', 'Reinsurance Premium']
option_menu=ttk.OptionMenu(win, option_control_variable, *options)
option_menu.grid(column=0, row=0)
option_menu.config(width=20)

ttk.Label(win, text="Chọn file BCTH05A", width=20).grid(column=0, row=1)
button1 = ttk.Button(win, text="BCTH05A", command=click_select_cpc)
button1.grid(column=1, row=1)

ttk.Label(win, text="Chọn file CTGS", width=20).grid(column=0, row=2)
button2 = ttk.Button(win, text="CTGS", command=click_select_acc)
button2.grid(column=1, row=2)

ttk.Label(win, text="Chọn thư mục lưu file", width=20).grid(column=0, row=3)
button2 = ttk.Button(win, text="Folder", command=click_save_file)
button2.grid(column=1, row=3)

ttk.Label(win, text="Enter saving file name", width=20).grid(column=0, row=4)
saving_file_name = ttk.Entry(win, width=10)
saving_file_name.grid(column=1, row=4)

ttk.Label(win, text="Xử lý số liệu", width=20).grid(column=0, row=5)
button3 = ttk.Button(win, text="Run", command=click_execute)
button3.grid(column=1, row=5)

label_open_saved_file = ttk.Label(win, text="Open saved file", width=30).grid(column=0, row=6)
button_open_saved_file = ttk.Button(win, text="Open saved file", command=button_open_saved_file_click)
button_open_saved_file.grid(column=1, row=6)

win.mainloop()
