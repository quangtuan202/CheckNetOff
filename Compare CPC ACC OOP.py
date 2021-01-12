#####################################################################
# Import libs required

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
    def __init__(self,file_name,use_cols,original_cols,new_cols,sum_cols,key):
        # use_cols : Columns from source file
        # original_cols : Columns from source file that must be renamed
        # new_cols : New name of Columns from source file that must be renamed
        # total_cols : Columns that must be added together
        # Create dataframe from file
        self.sum_cols=sum_cols
        self.key=key
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
        self.dataframe.rename(columns=dict(zip(original_cols,new_cols)), inplace=True)
    def new_policy_cols(self):
        # Create a columns for key with '.' removed
        self.dataframe[self.key+'_dot']=self.dataframe[self.key].apply(self.policy_dot)

        # Create a columns for key with '-' removed
        self.dataframe[self.key+'_hyphen']=self.dataframe[self.key].str.split('-').str[0]

        # Create a columns for key with 'EN' removed
        self.dataframe[self.key+'_EN']=self.dataframe[self.key].str.split('EN').str[0]

        # Create a columns for key contains character to dept
        self.dataframe[self.key+'_dept']=self.dataframe[self.key].apply(self.policy_dept)

#################################################################################

    def total_col(self):
        if len(self.sum_cols)!=0:
            self.dataframe['Total']=self.dataframe[self.sum_cols[0]].copy()
            for i in range(1,len(self.sum_cols)):
                self.dataframe['Total']=self.dataframe['Total']+self.dataframe[[i]]

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
    
##############################################################################
class Dataframe_cpc(Dataframe):
    def __init__(self,file_name,use_cols,original_cols,new_cols,sum_cols,key):
        super().__init__(file_name,use_cols,original_cols,new_cols,sum_cols,key)

##############################################################################
class Dataframe_acc(Dataframe):
    def __init__(self,file_name,use_cols,original_cols,new_cols,sum_cols,account_list,account_type,cols_of_accounts,key):
        super().__init__(file_name,use_cols,original_cols,new_cols,sum_cols,key)
        self.account_type=account_type
        self.account_list=account_list
        # Create a combined list of accounts from nested list
        combined_list=self.account_list[0]
        for i in range(1,len(self.account_list)):
            combined_list=combined_list+self.account_list[i]
        self.acc_combined_list=combined_list
        self.dataframe[key]=self.dataframe.apply(lambda x: return_policy(x['Số đơn'], x['Số đơn HT'], x['Số đơn ĐT'], x['TK Nợ'], x['TK Có']), axis=1)
        self.dataframe['NT_acc'] = self.dataframe.apply(lambda x: return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền']), axis=1)
        self.dataframe['VND_acc'] = self.dataframe.apply(lambda x: return_amount(x['TK Nợ'], x['TK Có'], x['Số tiền VND']), axis=1)

        for i in range(0,len(cols_of_accounts)):
            self.dataframe[cols_of_accounts[i]]=self.dataframe.loc[self.dataframe['TK Nợ'].isin([account_list[i]]) | self.dataframe['TK Nợ'].isin([account_list[i]]),['VND_acc']]
        super().total_col()
      

#-----------------------------------------------------------------------    
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


# --------------------------------------------------------------------------------

    def return_policy(self,don_ht, don_dt, don, debit_account, credit_account):
        if len(don_ht) < 2 and len(don_dt) < 2:
            return don
        elif len(don_ht) < 2 and len(don) < 2:
            return don_dt
        elif len(don_dt) < 2 and len(don) < 2:
            return don_ht
        elif debit_account in self.acc_combined_list:
            return don_dt
        elif credit_account in self.acc_combined_list:
            return don_ht
        else:
            return


# Inherit cpc class, extend to cover account codes
    


############################################################################
        conditions_dept=[self.dataframe[key].str.contains("SYCAR")
                        ,self.dataframe[key].str.contains("HY")
                        ,self.dataframe[key].str.contains("HN") & (~self.dataframe[key].str.contains(".HN"))
                        ,self.dataframe[key].str.contains("HL")
                        ,self.dataframe[key].str.contains("SY")
                        ,self.dataframe[key].str.contains("SN")
                        ,self.dataframe[key].str.contains("HB")
                        ,self.dataframe[key].str.contains("SB")
                        ,self.dataframe[key].str.contains("HP")
                        ,self.dataframe[key].str.contains("HR")
                        ,self.dataframe[key].str.contains("SR")
                        ,self.dataframe[key].str.contains("HG")
                        ,self.dataframe[key].str.contains("HU")
                        ,self.dataframe[key].str.contains("DR")
                        ,self.dataframe[key].str.contains("VR")
                        ,self.dataframe[key].str.contains("SG")]
        
        choices_dept=[self.dataframe[key].str.slice(0,self.dataframe[key].str.find("SYCAR")+5,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HY")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HN")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HL")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("SY")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("SN")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HB")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("SB")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HP")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HR")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("SR")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HG")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("HU")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("DR")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("VR")+3,1)
                    ,self.dataframe[key].str.slice(0,self.dataframe[key].str.find("SG")+3,1)]

        self.dataframe[key+'_Dept']=np.select(conditions_dept,choices_dept, default=self.dataframe[key])



premium_acc_list = [511111, 511112, 511113, 511114, 511115, 511116, 511131, 511133, 511134, 511136, 511211,
                    531111, 531112, 531113, 531114, 531115, 531116, 531131, 531133, 531134, 531136, 531211,
                    532111, 532112, 532113, 532114, 532115, 532116, 532117]


commission_acc_list = [624141,624143,624241,624173]
brokerage_acc_list= [624142,624144]
claim_acc_list = [624111,624118,624162,624211]
survey_fee=[624112]
claim_proceeds=[513811,513812]



        ##################################################################################################
        # Instance of window
win.mainloop()