import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox as msg
import pyxlsb
import pandas as pd
import numpy as np
import subprocess

####################################################################
# Initiate Tkinter window

win = tk.Tk()
win.title("UIC ACC-CPC Reconciliation")

#####################################################################
# Make a class for data frame:

class Dataframe:
    def __init__(self,file_name,use_cols,original_cols,new_cols,sum_cols,final_cols,key):
        # use_cols : Columns from source file
        # original_cols : Columns from source file that must be renamed
        # new_cols : New name of Columns from source file that must be renamed
        # total_cols : Columns that must be added together
      
        self.sum_cols=sum_cols
        self.final_cols=final_cols
        self.final_cols_dot=[]
        self.final_cols_hyphen=[]
        self.final_cols_en=[]
        self.final_cols_dept=[]
        self.key=key

        # Create dataframe from file
        if file_name.endswith('csv'):
            self.dataframe=pd.read_csv(file_name, usecols=use_cols)
        elif file_name.endswith('xlsb'):
            self.dataframe = pd.read_excel(file_name, usecols=use_cols, engine='pyxlsb')
        elif file_name.endswith('xls') or file_name.endswith('xlsx'):
            self.dataframe = pd.read_excel(file_name, usecols=use_cols)
        else: 
            msg.showinfo("Information", "Select Excel or CSV file only")
            pass
        

        # Rename columns
    #-------------------------------------------------------------------------------------------     
    def combine_list(self,account_list):
        try:
            acc_combined_list=[]
            for i in range(0,len(account_list)):
                for j in range(0,len(account_list[i])):
                    acc_combined_list.append(account_list[i][j])
        except:
            acc_combined_list=account_list
        return acc_combined_list  

    #------------------------------------------------------------------------------------------- 
    def new_key(self):
        # Create a columns for key with '.' removed
        self.dataframe[self.key+'_dot']=self.dataframe[self.key].apply(self.policy_dot)

        # Create a columns for key with '-' removed
        self.dataframe[self.key+'_hyphen']=self.dataframe[self.key].str.split('-').str[0]

        # Create a columns for key with 'EN' removed
        self.dataframe[self.key+'_en']=self.dataframe[self.key].str.split('EN').str[0]

        # Create a columns for key contains character to dept
        self.dataframe[self.key+'_dept']=self.dataframe[self.key].apply(self.policy_dept)
        
        # Create a columns for dept name
        #self.dataframe['Dept']=self.dataframe[self.key].apply(self.policy_dept)
    #-------------------------------------------------------------------------------------------        
    def sub_dataframe(self):       
            
        self.dataframe=self.dataframe[self.final_cols].copy()
        self.dataframe = self.dataframe.groupby(key, as_index=False).sum()
        return self.dataframe
    #-------------------------------------------------------------------------------------------
    def sub_dataframe_dot(self):

        self.final_cols_dot=self.final_cols.copy()
        self.final_cols_dot[0]=self.final_cols_dot[0]+'_dot' 

        self.dataframe_dot = self.dataframe[self.final_cols_dot].copy()
        self.dataframe_dot = self.dataframe_dot.groupby(self.key+'_dot', as_index=False).sum()  # as_index=False to show grouped column
        new_name_dot= [x+'dot' for x in final_cols] # create a list of final_cols + dot
        self.dataframe_dot.rename(columns=dict(zip(final_cols,new_name_dot)), inplace=True)
        return self.dataframe_dot
    #---------------------------------------------------------------------------------------------
    def sub_dataframe_hyphen(self):

        self.final_cols_hyphen=self.final_cols.copy()
        self.final_cols_hyphen[0]=self.final_cols_hyphen[0]+'_hyphen' 

        self.dataframe_hyphen = self.dataframe[self.final_cols_hyphen].copy()
        self.dataframe_hyphen = self.dataframe_hyphen.groupby(self.key+'_hyphen', as_index=False).sum()  # as_index=False to show grouped column
        new_name_hyphen= [x+'hyphen' for x in final_cols] # create a list of final_cols + hyphen
        self.dataframe_hyphen.rename(columns=dict(zip(final_cols,new_name_hyphen)), inplace=True)
        return self.dataframe_hyphen
    #--------------------------------------------------------------------------------------------
    def sub_dataframe_en(self):

        self.final_cols_en=self.final_cols.copy()
        self.final_cols_en[0]=self.final_cols_en[0]+'_en' 

        self.dataframe_en = self.dataframe[self.final_cols_en].copy()
        self.dataframe_en = self.dataframe_en.groupby(self.key+'_en', as_index=False).sum()  # as_index=False to show grouped column
        new_name_en= [x+'en' for x in final_cols] # create a list of final_cols + en
        self.dataframe_en.rename(columns=dict(zip(final_cols,new_name_en)), inplace=True)
        return self.dataframe_en
    #--------------------------------------------------------------------------------------------
    def sub_dataframe_dept(self):

        self.final_cols_dept=self.final_cols.copy()
        self.final_cols_dept[0]=self.final_cols_dept[0]+'_dept' 

        self.dataframe_dept = self.dataframe[self.final_cols_dept].copy()
        self.dataframe_dept = self.dataframe_dept.groupby(self.key+'_dept', as_index=False).sum()  # as_index=False to show grouped column
        new_name_dept= [x+'dept' for x in final_cols] # create a list of final_cols + dept
        self.dataframe_dept.rdeptame(columns=dict(zip(final_cols,new_name_dept)), inplace=True)
        return self.dataframe_dept

    #--------------------------------------------------------------------------------------------

    def total_col(self):
        if len(self.sum_cols)!=0:
            self.dataframe['Total']=self.dataframe[self.sum_cols[0]].copy()
            for i in range(1,len(self.sum_cols)):
                self.dataframe['Total']=self.dataframe['Total']+self.dataframe[self.sum_cols[i]]

    #--------------------------------------------------------------------------------------------
    
    def policy_dot(self,policy):
        if policy.startswith('00') and policy.endswith('T') and policy.find(".") == -1:
            return policy[:-1]
        elif policy.startswith('ARRTRIP.'):
            a=len(policy)+1
            return policy[8:a]
        elif policy.find(".") != -1:
            return policy[0:policy.find(".")]
        else:
            return policy

    #--------------------------------------------------------------------------------------------
    
    def policy_dept(self,policy):
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
    #--------------------------------------------------------------------------------------------
    def return_dept_name(self,codename):
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



##############################################################################
class Dataframe_cpc(Dataframe):
    def __init__(self,file_name,use_cols,original_cols,new_cols,sum_cols,final_cols,key):
        super().__init__(file_name,use_cols,original_cols,new_cols,sum_cols,final_cols,key)

##############################################################################
class Dataframe_acc(Dataframe):
    def __init__(self,file_name,use_cols,original_cols,new_cols,sum_cols,account_list,account_type,cols_of_accounts,final_cols,key):
        super().__init__(file_name,use_cols,original_cols,new_cols,sum_cols,final_cols,key)
        self.account_type=account_type
        self.account_list=account_list
        self.dataframe[['Số đơn', 'Số đơn HT', 'Số đơn ĐT']]=self.dataframe[['Số đơn', 'Số đơn HT', 'Số đơn ĐT']].fillna('0')
        # Create a combined list of accounts from nested list
        self.acc_combined_list=self.combine_list(self.account_list)
 
        self.dataframe[self.key]=self.dataframe.apply(lambda x: self.return_policy(x['Số đơn'], x['Số đơn HT'], x['Số đơn ĐT'], x['TK Nợ'], x['TK Có']), axis=1)
        #self.key=self.dataframe[key].name
        self.dataframe['NT'] = self.dataframe.apply(lambda x: self.return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền']), axis=1)
        self.dataframe['VND'] = self.dataframe.apply(lambda x: self.return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền VND']), axis=1)
        self.dataframe = self.dataframe[~pd.isnull(self.dataframe['VND'])]
      
        try:
            for i in range(0,len(cols_of_accounts)):
                self.dataframe[cols_of_accounts[i]]=self.dataframe.loc[(self.dataframe['TK Nợ'].isin(account_list[i]) | self.dataframe['TK Có'].isin(account_list[i])),['VND']]
                self.dataframe[cols_of_accounts[i]]=self.dataframe[cols_of_accounts[i]].fillna(0)
        except:
            pass
                
###################################################################################    
  
###############################################################################  
    def return_policy(self,don,don_ht, don_dt,debit_account, credit_account):
        if len(don_ht) < 8 and len(don_dt) < 8:
            return don
        elif len(don_ht) < 8 and len(don) < 8:
            return don_dt
        elif len(don_dt) < 8 and len(don) < 8:
            return don_ht
        elif debit_account in self.acc_combined_list:
            return don_dt
        elif credit_account in self.acc_combined_list:
            return don_ht
        else:
            return
    ###############################################################################
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

#########################################################################################
# Class for list of unique policies object
class list_of_unique_item():
    def __init__(self,list_a,list_b,key):
        self.key=key
        self.df=pd.concat([list_a,list_b]).unique()
        self.dataframe=pd.DataFrame(self.df,columns=[key])

#########################################################################################
# Class for joining dataframe objects
class join_dataframe():
    def __init__(self,df_unique,acc_df,acc_df_dot,acc_df_hyphen,acc_df_en,acc_df_dept,cpc_df,cpc_df_dot,cpc_df_hyphen,cpc_df_en,cpc_df_dept,acc_df_name,cpc_df_name,cols_to_compare,cols_to_find_min,final_cols):
        # df_unique : list of unique policies
        # df_dot : dataframe with dot removed
        # df_hyphen : dataframe with hyphen removed
        # df_en : dataframe with en removed
        # df_dept : dataframe contains strings to dept only
        # cols_to_compare : pairs of columns to be compared , data type : dictionary {'NT':[col1,col2],'VND':[col1,col2]}
        # cols_to_find_min : list of columns to find min value. data type : dictionary {'NT':[col1,col2],'VND':[col1,col2]}
        self.df_merged=df_unique.merge(acc_df,how='left',on='Policy',suffixes=('_acc', '_cpc')).merge(cpc_df,how='left',on='Policy',suffixes=('_acc', '_cpc')).merge(acc_df_dot,how='left'
    ,on='Policy_dot',suffixes=('_acc', '_cpc')).merge(cpc_df_dot,how='left',on='Policy_dot',suffixes=('_acc', '_cpc')).merge(acc_df_hyphen,how='left',on='Policy_hyphen',suffixes=('_acc', '_cpc')).merge(cpc_df_hyphen,how='left'
    ,on='Policy_hyphen',suffixes=('_acc', '_cpc')).merge(acc_df_en,how='left',on='Policy_en',suffixes=('_acc', '_cpc')).merge(cpc_df_en,how='left',on='Policy_en',suffixes=('_acc', '_cpc')).merge(acc_df_dept,how='left',on='Policy_dept',suffixes=('_acc', '_cpc')).merge(cpc_df_dept,how='left',on='Policy_dept',suffixes=('_acc', '_cpc')).merge(acc_df_name,how='left',on='Policy').merge(cpc_df_name,how='left',on='Policy')  
        
        values={'NT_cpc':0,'NT_cpc_dot':0,'NT_cpc_hyphen':0,'NT_cpc_en':0,'NT_cpc_dept':0,'VND_cpc':0,'VND_cpc_dot':0,
           'VND_cpc_hyphen':0,'VND_cpc_en':0,'VND_cpc_dept':0,
           'NT_acc':0,'NT_acc_dot':0,'NT_acc_hyphen':0,'NT_acc_en':0,'NT_acc_dot_hyphen_en':0,
           'NT_acc_dept':0,'VND_acc':0,'VND_acc_dot':0,'VND_acc_hyphen':0,'VND_acc_en':0,'VND_acc_dot_hyphen_en':0,'VND_acc_dept':0}
        self.df_merged.fillna(value=values,inplace=True)

        for key, value in cols_to_compare.items():
            self.df_merged['Chenh lech '+key]=self.df_merged[value[0]]-self.df_merged[value[1]]
            self.df_merged['Chenh lech '+key+'_abs']=self.df_merged['Chenh lech '+key].abs()

        for key, value in cols_to_find_min.items():
            self.df_merged['Min diff'+key]=self.df_merged.apply(lambda x: min(x[value]),axis=1) # value is a list

        self.df_merged['Client name']=np.where(self.df_merged['acc_client_name'].isnull(),self.df_merged['cpc_client_name'],self.df_merged['acc_client_name'])
        self.df_final=self.df_merged[final_cols].copy()
        


#***************************************************************Codes for GUI***************************************************************************

####################################################################
# Master data-accounts
gross_premium_accounts=[511111, 511112, 511113, 511114, 511115, 511116, 511131, 511133, 511134, 511136, 511211,
                    531111, 531112, 531113, 531114, 531115, 531116, 531131, 531133, 531134, 531136, 531211,
                    532111, 532112, 532113, 532114, 532115, 532116, 532117]
ri_premium_accounts=[533111,511411]
gross_claim_accounts=[624111,624118,624162,624211]
salvage_claim_accounts=[513811,513812]
survey_fee_accounts=[624112]
ri_claim_accounts=[624161]
claim_accounts_list=[gross_claim_accounts,survey_fee_accounts,salvage_claim_accounts]
commission_accounts=[624141,624143,624241,624173]
brokerage_accounts=[624142,624144]
commission_accounts_list=[commission_accounts,brokerage_accounts]
tpa_acc_list = [624181]

# Master data-use_cols
use_cols_premium_cpc=['Số đơn/ Số Endor','Tên Khách hàng', 'Phí Bảo Hiểm\n(nguyên tệ)','Phí bảo hiểm\n(VND)']
use_cols_commission_cpc=['Số đơn/ Số Endor','Môi giới phí BHG / Hoa hồng NTBH','Hoa hồng đại lý bảo hiểm']
use_cols_claim_cpc=['Số hồ sơ bồi thường\t','Tên Khách hàng','Bồi thường\n(VND)','Phí giám định']
use_cols_ri_claim_cpc=['Số hồ sơ bồi thường\t','Tên Khách hàng','Thu từ nhượng TBH FAC']
use_cols_tpa_cpc=['Số đơn/ Số Endor','Tên Khách hàng','Số tiền TPA']
use_cols_reinsurance_cpc=['Số đơn/ Số Endor','Tên Khách hàng','Phí nhượng TBH FAC']
use_cols_acc=['Số đơn', 'Số đơn HT', 'Số đơn ĐT','Tên ĐT giao dịch','TK Nợ', 'TK Có', 'Số tiền', 'Số tiền VND']
use_cols_cpc_name_of_client_claim=['Số hồ sơ bồi thường\t','Tên Khách hàng']
# Value and columns to fill Nan
values_fillna_gwp={'NT_cpc':0,'NT_dot_cpc':0,'NT_hyphen_cpc':0,'NT_en_cpc':0,'NT_dept_cpc':0,'VND_cpc':0,'VND_dot_cpc':0,
           'VND_hyphen_cpc':0,'VND_en_cpc':0,'VND_dept_cpc':0,
           'NT_acc':0,'NT_dot_acc':0,'NT_hyphen_acc':0,'NT_en_acc':0,
           'NT_dept_acc':0,'VND_acc':0,'VND_dot_acc':0,'VND_hyphen_acc':0,'VND_en_acc':0,'VND_dept_acc':0}
# đang code dở, cần thêm các value khác để fill Nan cho các case khác
values_fillna_claim

# -------------------------------------------------------------------------------------------------------------------
# Variables for file name and folder
cpc_file=''
acc_file=''
saving_folder=''
#**********************************************************Functions***********************************************************************
def return_gwp_df():
    #file_name,use_cols,original_cols,new_cols,sum_cols,account_list,account_type,cols_of_accounts,final_cols,key)
    file_name=acc_file
    use_cols=use_cols_acc
    original_cols=use_cols_acc
    new_cols=use_cols_acc
    sum_cols=[] # no total col
    account_list=gross_premium_accounts
    account_type='credit'
    cols_of_accounts=[]
    final_cols=['Policy','NT','VND']
    key='Policy'
    # Create acc instance 
    df_acc=Dataframe_acc(file_name,use_cols,original_cols,new_cols,sum_cols,account_list,account_type,cols_of_accounts,final_cols,key)
    df_acc.new_key()
    sub_df_acc=df_acc.sub_dataframe()
    sub_df_acc_dot=df_acc.sub_dataframe_dot()
    sub_df_acc_hyphen=df_acc.sub_dataframe_hyphen()
    sub_df_acc_en=df_acc.sub_dataframe_en()
    sub_df_acc_dept=df_acc.sub_dataframe_dept()
    # Create cpc instance 
    file_name=cpc_file
    use_cols=use_cols_premium_cpc
    original_cols=use_cols_premium_cpc
    new_cols=['Policy','NT','VND']
    sum_cols=[] # no total col
    account_list=gross_premium_accounts
    account_type='credit'
    cols_of_accounts=[]
    final_cols=['Policy','NT','VND']
    key='Policy'

    df_cpc=Dataframe_cpc(file_name,use_cols,original_cols,new_cols,sum_cols,final_cols,key)
    df_cpc.new_key()
    sub_df_cpc=df_cpc.sub_dataframe()
    sub_df_cpc_dot=df_cpc.sub_dataframe_dot()
    sub_df_cpc_hyphen=df_cpc.sub_dataframe_hyphen()
    sub_df_cpc_en=df_cpc.sub_dataframe_en()
    sub_df_cpc_dept=df_cpc.sub_dataframe_dept()

 
#**********************************************************GUI commands********************************************************************
def click_select_cpc():
    global cpc_file
    cpc_file = fd.askopenfilename()
#-------------------------------------------------

def click_select_acc():
    global acc_file
    acc_file = fd.askopenfilename()
#------------------------------------------------

def click_save_file():
    global saving_folder
    saving_folder = fd.askdirectory()
    #print(saving_folder)
#------------------------------------------------

def button_open_saved_file_click():
    global cpc_file
    global acc_file
    global saving_folder
    subprocess.Popen([f"{saving_folder}/{saving_file_name.get()}.csv"], shell=True)
    #win.destroy()
    cpc_file=''
    acc_file=''
    saving_folder=''
#----------------------------------------------

def click_execute():
    # Create DataFrame_acc object ==> Call New_key method ==> Call sub_dataframe , dot, hyphen, en, dept method : get 5 dataframes for ACC
    # Create DataFrame_cpc object ==> Call New_key method ==> Call sub_dataframe , dot, hyphen, en, dept method : get 5 dataframes for CPC
    # Create a list that combines ACC & CPC policies, remove duplicate policies
    # Create a list of policy and client name from ACC and CPC
    # Join the list of unique item and the above dataframe

    from tkinter import messagebox as msg
    # "Revenue", "Commission", "Claim", 'TPA', 'Reinsurance Premium'
    if option_control_variable.get()=='Gross Written Premium':
        use_cols_cpc=use_cols_premium_cpc
    elif option_control_variable.get()=='Commission':
        use_cols_cpc=use_cols_commission_cpc
    elif option_control_variable.get()=='Claim':
        use_cols_cpc=use_cols_claim_cpc
    elif option_control_variable.get()=='TPA':
        use_cols_cpc=use_cols_tpa_cpc
    elif option_control_variable.get()=='Reinsurance Premium':
        use_cols_cpc=use_cols_


# GUI elements------------------------------------------------------------------
option_control_variable = tk.StringVar()
options = ['Select',"Gross Written Premium", "Commission", "Claim", 'TPA', 'Reinsurance Premium']
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
    
