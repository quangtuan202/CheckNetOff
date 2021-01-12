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

####################################################################
# Make a class for data frame:

class Dataframe:
    def __init__(self,file_name,use_cols,original_cols,new_cols,sum_cols,final_cols,key):
        # use_cols : Columns from source file
        # original_cols : Columns from source file that must be renamed
        # new_cols : New name of Columns from source file that must be renamed
        # total_cols : Columns that must be added together
        #final_cols :
        # key: policy or claim number 
      
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
    def sub_dataframe(self):
        # Create a columns for key with '.' removed
        self.dataframe[self.key+'_dot']=self.dataframe[self.key].apply(self.policy_dot)

        # Create a columns for key with '-' removed
        self.dataframe[self.key+'_hyphen']=self.dataframe[self.key].str.split('-').str[0]

        # Create a columns for key with 'EN' removed
        self.dataframe[self.key+'_en']=self.dataframe[self.key].str.split('EN').str[0]

        # Create a columns for key contains character to dept
        self.dataframe[self.key+'_dept']=self.dataframe[self.key].apply(self.policy_dept)
        
        self.dataframe = self.dataframe.groupby(self.key, as_index=False).sum()

        # Create a list of final cols to select
     
        self.dataframe_dot = self.dataframe[self.final_cols].copy()
        self.dataframe_dot = self.dataframe_dot.groupby(self.key, as_index=False).sum()  # as_index=False to show grouped column
        new_name_dot= [x+'dot' for x in final_cols] # create a list of final_cols + dot
        self.dataframe_dot.rename(columns=dict(zip(final_cols,new_name_dot)), inplace=True)

        self.dataframe_hyphen = self.dataframe[self.final_cols].copy()
        self.dataframe_hyphen = self.dataframe_hyphen.groupby(self.key, as_index=False).sum()  # as_index=False to show grouped column
        new_name_hyphen= [x+'hyphen' for x in final_cols] # create a list of final_cols + hyphen
        self.dataframe_hyphen.rename(columns=dict(zip(final_cols,new_name_hyphen)), inplace=True)

        self.dataframe_en = self.dataframe[self.final_cols].copy()
        self.dataframe_en = self.dataframe_en.groupby(self.key, as_index=False).sum()  # as_index=False to show grouped column
        new_name_en= [x+'en' for x in final_cols] # create a list of final_cols + en
        self.dataframe_en.rename(columns=dict(zip(final_cols,new_name_en)), inplace=True)


        self.dataframe_dept = self.dataframe[self.final_cols].copy()
        self.dataframe_dept = self.dataframe_dept.groupby(self.key, as_index=False).sum()  # as_index=False to show grouped column
        new_name_dept= [x+'dept' for x in final_cols] # create a list of final_cols + dept
        self.dataframe_dept.rename(columns=dict(zip(final_cols,new_name_dept)), inplace=True)
            
        return self.dataframe_dot
        return self.dataframe_hyphen
        return self.dataframe_en
        return self.dataframe_dept

#################################################################################

    def total_col(self):
        if len(self.sum_cols)!=0:
            self.dataframe['Total']=self.dataframe[self.sum_cols[0]].copy()
            for i in range(1,len(self.sum_cols)):
                self.dataframe['Total']=self.dataframe['Total']+self.dataframe[self.sum_cols[i]]

###################################################################################
    
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

####################################################################################
    
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
###############################################################################


    
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
        self.acc_combined_list=[]
        for i in range(0,len(self.account_list)):
            self.acc_combined_list.append(self.account_list[i])
            
        self.dataframe[key]=self.dataframe.apply(lambda x: self.return_policy(x['Số đơn'], x['Số đơn HT'], x['Số đơn ĐT'], x['TK Nợ'], x['TK Có']), axis=1)
        self.key=self.dataframe[key].name
        self.dataframe['NT'] = self.dataframe.apply(lambda x: self.return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền']), axis=1)
        self.dataframe['VND'] = self.dataframe.apply(lambda x: self.return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền VND']), axis=1)
        self.dataframe = self.dataframe[~pd.isnull(self.dataframe['VND'])]
      
        for i in range(0,len(cols_of_accounts)):
            self.dataframe[cols_of_accounts[i]]=self.dataframe.loc[(self.dataframe['TK Nợ'].isin(account_list[i]) | self.dataframe['TK Có'].isin(account_list[i])),['VND']]
            self.dataframe[cols_of_accounts[i]]=self.dataframe[cols_of_accounts[i]].fillna(0)
        #super().total_col()
        #if len(cols_of_accounts)!=0:
            #self.dataframe['Total']=self.dataframe[cols_of_accounts[0]].copy()
            #for i in range(1,len(cols_of_accounts)):
                #self.dataframe['Total']=self.dataframe['Total']+self.dataframe[cols_of_accounts[i]]
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