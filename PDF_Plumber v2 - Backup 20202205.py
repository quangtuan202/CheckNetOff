import pdfplumber
import pandas as pd
import numpy as np
import os
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog as fd
import subprocess
from pandas import ExcelWriter

# ---------------------------------
win = tk.Tk()
win.title("PDF conversion")
win.geometry('510x200')
win.iconbitmap('D:/OneDrive\OneDrive - khoavanhoc.edu.vn/Python Project/pdf_icon_55_UOU_icon.ico')

file_path = ''
folder_path=''
premium_or_claim = ''
saving_folder_path = ''
line_or_text_method = ''
file_or_folder=''
df = []
fx_rate=[]
currency=[]
premium_columns = []
claim_columns = []
df_concat = pd.DataFrame()
df_concat_1=pd.DataFrame()
df_concat_tmp=pd.DataFrame()
fx_table=pd.DataFrame()

lob={'IAR':'5','HIO':'7','HAS':'a.3','PAM':'a.1','PAI':'a.1','WCP':'a.1','WCI':'a.1','TRA':'a.1','TRF':'a.1','PUL':'7','PRL':'7','PAR':'5','HMR':'1','NES':'7','TFT':'1'
,'MON':'1','OMR':'1','CAR':'1','EAR':'1','MBD':'1','EEI':'1','FIR':'5','TRD':'a.1','REL':'7','BBB':'8','MDI':'a.3','CER':'1','SDI':'1','LAR':'1','CPM':'1','FGI':'7','PPL':'7'
,'MCI':'2','MCE':'2','ICA':'2','HCI':'a.3','HUL':'6','MOX':'4','MOT':'4','TOL':'7','P&I':'6','CPH':'a.3','MDL':'a.3','MPA':'1','CRE':'8','RPC':'1','PII':'7','BLI':'7','BFL':'7'
,'COT':'2','BPV':'1','CLI':'6','CGL':'7','FFL':'7','IPH':'a.3','GPH':'a.3','KID':'a.3','FSI':'a.3','TRR':'1','GOL':'7','CPA':'a.3','ACA':'2','MCA':'2','AGI':'10','ECI':'a.3'
,'000':'4','001':'4','002':'4','BII':'5','DNO':'7','PNI':'6','WRH':'6','ENG':'1','PKI':'5','KPI':'5','HOM':'5','ENG':'1','CTP':'4'}  

lob_name={'IAR':'Fire and explosion','HIO':'Public liability','HAS':'Health','PAM':'PA','PAI':'PA','WCP':'PA','WCI':'PA','TRA':'TRA','TRF':'TRA','PUL':'Public liability'
,'PRL':'Public liability','PAR':'Fire and explosion','HMR':'Eng','NES':'Public liability','TFT':'Eng','MON':'Eng','OMR':'Eng','CAR':'Eng','EAR':'Eng','MBD':'Eng','EEI':'Eng'
,'FIR':'Fire and explosion','TRD':'TRA','REL':'Public liability','BBB':'Credit insurance ','MDI':'Health','CER':'Eng','SDI':'Eng','LAR':'Eng','CPM':'Eng','FGI':'Public liability'
,'PPL':'Public liability','MCI':'Cargo Insurance','MCE':'Cargo Insurance','ICA':'Cargo Insurance','HCI':'Health','HUL':'Hull insurance ','MOX':'Motor insurance','MOT':'Motor insurance'
,'TOL':'Public liability','P&I':'Hull insurance ','CPH':'Health','MDL':'Health','MPA':'Eng','CRE':'Credit insurance ','RPC':'Eng','PII':'Public liability','BLI':'Public liability'
,'BFL':'Public liability','COT':'Cargo Insurance','BPV':'Eng','CLI':'Hull insurance ','CGL':'Public liability','FFL':'Public liability','IPH':'Health','GPH':'Health','KID':'Health'
,'FSI':'Health','TRR':'Eng','GOL':'Public liability','CPA':'Health','ACA':'Cargo Insurance','MCA':'Cargo Insurance','AGI':'Agricultural','ECI':'Health','000':'Motor insurance'
,'001':'Motor insurance','002':'Motor insurance','BII':'Fire and explosion','DNO':'Public liability','PNI':'Hull insurance ','WRH':'Hull insurance ','ENG':'Eng','PKI':'Fire and explosion'
,'KPI':'Fire and explosion','HOM':'Fire and explosion','CTP':'Motor insurance'}

#----------------------------------------------------return department name

def get_dept_name(codename):
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
        return 'RI'
    elif 'HUD' in codename:
        return 'RI'
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
        return 'Retail (HCMC)'
    elif codename[10:12] == '41':
        return 'Retail (HN)'
    elif codename[10:12] == '42':
        return 'Retail (HCMC)'
    elif codename[10:12] == '01':
        return 'RI'
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

#-----------------------------------------------------return policy type

def get_policy_type(policy):
    if '501' in policy and policy.startswith('00'):
        return 'MOT'
    elif '502' in policy and policy.startswith('00'):
        return 'CTP'
    elif 'CTP' in policy:
        return 'CTP'
    elif 'MOT' in policy and policy.startswith('0'):
        return 'MOT'
    elif 'MOX' in policy and policy.startswith('0'):
        return 'MOX'
    elif 'ARRTRIP' in policy:
        return policy[8:11]
    else:
        return policy[0:3]

#----------------------------------------------------Return policy LOB

def get_lob(policy_type):
    if policy_type in lob.keys():
        return lob.get(policy_type)

#-----------------------------------------------------Return policy LOB name

def get_lob_name(policy_type):
    if policy_type in lob_name.keys():
        return lob_name.get(policy_type)      

#-----------------------------------------------------Remove duplicate FX

def remove_fx_duplicate(tbl=pd.DataFrame()):
    col_1=list(tbl['FX rate'])
    col_2=list(tbl['Currency'])
    col_3=list(tbl['File_name'])
    col_1_new=[]
    col_2_new=[]
    col_3_new=[]
    col_4_new=[]
    for i in range(0,len(col_1)):
        if col_1[i-1]!=col_1[i] or col_2[i-1]!=col_2[i] or col_3[i-1]!=col_3[i]:
            col_1_new.append(col_1[i])
            col_2_new.append(col_2[i])
            col_3_new.append(col_3[i])
    col_4_new=[ x for x in range(len(col_1))]
    df=pd.DataFrame(list(zip(col_1_new,col_2_new,col_3_new,col_4_new)),columns=['FX rate','Currency','file_name','fx_order_num']) 
    return df

#-----------------------------------------------------Create Fx_order_number for df_concat_1

def set_fx_order_num(tbl=pd.DataFrame()):
    a= list(tbl['Currency'].copy())
    b= list(tbl['File name'].copy())
    c= [ 0 for x in range(len(a))]
    for i in range(1,len(a)):
        if a[i]==a[i-1] and b[i]==b[i-1]:
            c[i]=c[i-1]
        else:
            c[i]=c[i-1]+1
    return c


#-----------------------------------------------------Read a single file only

def read_file():
    config = {"vertical_strategy": "lines", "horizontal_strategy": "text", }

    global line_or_text_method
    global df
    global fx_rate
    global currency
    global fx_table
    #words=[]
    file_name=[]
    pdf = pdfplumber.open(file_path)
    for i in range(len(pdf.pages)):
        if line_or_text_method == 'text':
            table = pdf.pages[i].extract_tables(config)
            for j in range(len(table)):
                dataframe = pd.DataFrame(table[j])
                dataframe['File name'] = file_path
                df.append(dataframe)
        else:
            table = pdf.pages[i].extract_tables()
            for j in range(len(table)):
                dataframe = pd.DataFrame(table[j])
                dataframe['File name'] = file_path
                df.append(dataframe)
        words=pdf.pages[i].extract_words()
                #words.append(word_per_page)
        for x in range(len(words)):
            if words[x]['text']=='EXCHANGE':
                if words[x+3]['text']=='EQUAL':
                    currency.append(words[x+5]['text'])
                    fx_rate.append(words[x+2]['text'])
                    file_name.append(file_path)
                elif words[x+4]['text']=='EQUAL':
                    currency.append(words[x+6]['text'])
                    fx_rate.append(words[x+2]['text']+words[x+3]['text'])
                    file_name.append(file_path)
    fx_table_temp=pd.DataFrame(list(zip(fx_rate,currency,file_name)),columns=['FX rate','Currency','File_name']) # chỗ này cần code tương tự như dataframe, tạo dataframe từ từng file rồi concat
    fx_table=remove_fx_duplicate(fx_table_temp)

    return df
    return fx_table

#--------------------------------------------------Read all files in one folder

def read_folder():
    config = {"vertical_strategy": "lines", "horizontal_strategy": "text", }

    global line_or_text_method
    global df
    global fx_rate
    global currency
    global fx_table
    #words=[]
    file_name=[]


    for filename in os.listdir(folder_path):
        name, ext = os.path.splitext(filename)
        if ext == '.pdf':
            pdf = pdfplumber.open(f"{folder_path}/{filename}")
            for i in range(len(pdf.pages)):
                if line_or_text_method == 'text':
                    table = pdf.pages[i].extract_tables(config)
                    for j in range(len(table)):
                        dataframe = pd.DataFrame(table[j])
                        dataframe['File name'] = filename
                        df.append(dataframe)
                else:
                    table = pdf.pages[i].extract_tables()
                    for j in range(len(table)):
                        dataframe = pd.DataFrame(table[j])
                        dataframe['File name'] = filename
                        df.append(dataframe)
                words=pdf.pages[i].extract_words()
                #words.append(word_per_page)
                for x in range(len(words)):
                    if words[x]['text']=='EXCHANGE':
                        if words[x+3]['text']=='EQUAL':
                            currency.append(words[x+5]['text'])
                            fx_rate.append(words[x+2]['text'])
                            file_name.append(filename)
                        elif words[x+4]['text']=='EQUAL':
                            currency.append(words[x+6]['text'])
                            fx_rate.append(words[x+2]['text']+words[x+3]['text'])
                            file_name.append(filename)
    fx_table_temp=pd.DataFrame(list(zip(fx_rate,currency,file_name)),columns=['FX rate','Currency','File_name']) # chỗ này cần code tương tự như dataframe, tạo dataframe từ từng file rồi concat
    fx_table=remove_fx_duplicate(fx_table_temp)
    #fx_table['File_name']=filename
    #fx_table['table_number']=fx_table.index

    return df
    return fx_table        

#--------------------------------------Get dataframe from premium files

def get_dataframe_premium(dtf):
    global line_or_text_method

    if line_or_text_method == 'text':
        if dtf.shape[1] == 17:
            dtf = dtf[[0, 1, 3, 4, 6, 7, 11, 12, 13, 14, 'File name']]
            dtf = dtf.rename(
                columns={0: 'Policy', 1: 'Insured', 3: 'From', 4: 'To', 6: 'Premium', 7: 'Currency', 11: 'RI Premium',
                         12: 'RI Comm', 13: 'RI Tax', 14: 'Net RI Premium'})
            dtf1 = dtf.loc[:,
                   ['Policy', 'Insured', 'From', 'To', 'Premium', 'Currency', 'RI Premium', 'RI Comm', 'RI Tax',
                    'Net RI Premium', 'File name']]
            dtf1.insert(1, 'Endorsement No', 'NA')
            return dtf1

        elif dtf.shape[1] == 19:
            dtf = dtf[[0, 1, 2, 5, 7, 8, 14, 15, 16, 17, 'File name']]
            dtf = dtf.rename(
                columns={0: 'Policy', 1: 'Endorsement No', 2: 'Insured', 5: 'From', 7: 'Premium', 8: 'Currency',
                         14: 'RI Premium', 15: 'RI Comm', 16: 'RI Tax', 17: 'Net RI Premium'})
            dtf1 = dtf.loc[:,
                   ['Policy', 'Endorsement No', 'Insured', 'From', 'Premium', 'Currency', 'RI Premium', 'RI Comm',
                    'RI Tax', 'Net RI Premium', 'File name']]
            dtf1.insert(4, 'To', '01/01/1900')
            return dtf1

        elif dtf.shape[1] == 20:

            dtf = dtf[[0, 1, 2, 4, 5, 7, 8, 14, 15, 16, 17, 'File name']]
            dtf = dtf.rename(
                columns={0: 'Policy', 1: 'Endorsement No', 2: 'Insured', 4: 'From', 5: 'To', 7: 'Premium',
                         8: 'Currency', 14: 'RI Premium', 15: 'RI Comm', 16: 'RI Tax', 17: 'Net RI Premium'})
            dtf1 = dtf.loc[:,
                   ['Policy', 'Endorsement No', 'Insured', 'From', 'To', 'Premium', 'Currency', 'RI Premium', 'RI Comm',
                    'RI Tax', 'Net RI Premium', 'File name']]
            return dtf1

        else:
            return dtf

    else:
        if dtf.shape[1] == 21:
            dtf = dtf[[0, 1, 2, 3, 6, 8, 9, 15, 16, 17, 18, 'File name']]
            dtf = dtf.rename(
                columns={0: 'No', 1: 'Policy', 2: 'Endorsement No', 3: 'Insured', 6: 'From', 8: 'Premium',
                         9: 'Currency', 15: 'RI Premium', 16: 'RI Comm', 17: 'RI Tax', 18: 'Net RI Premium'})
            dtf1 = dtf.loc[:,
                   ['No', 'Policy', 'Endorsement No', 'Insured', 'From', 'Premium', 'Currency', 'RI Premium', 'RI Comm',
                    'RI Tax', 'Net RI Premium', 'File name']]
            dtf1.insert(5, 'To', '01/01/1900')
            return dtf1

        elif dtf.shape[1] == 19:
            dtf = dtf[[0, 1, 2, 4, 5,7, 8, 12, 13, 14, 15, 'File name']]
            dtf = dtf.rename(
                columns={0: 'No', 1: 'Policy', 2: 'Insured', 4: 'From', 5: 'To',7:'Premium', 8: 'Currency', 12: 'RI Premium',
                         13: 'RI Comm',
                         14: 'RI Tax', 15: 'Net RI Premium'})
            dtf1 = dtf.loc[:,
                   ['No', 'Policy', 'Insured', 'From', 'To', 'Currency', 'Premium','RI Premium', 'RI Comm', 'RI Tax',
                    'Net RI Premium',
                    'File name']]
            dtf1.insert(2, 'Endorsement No', 'NA')
            return dtf1

        elif dtf.shape[1] == 22:
            dtf = dtf[[0, 1, 2, 3, 5, 6, 8,9, 15, 16, 17, 18, 'File name']]
            dtf = dtf.rename(
                columns={0: 'No', 1: 'Policy', 2: 'Endorsement No', 3: 'Insured', 5: 'From', 6: 'To',8:'Premium', 9: 'Currency',
                         15: 'RI Premium', 16: 'RI Comm', 17: 'RI Tax', 18: 'Net RI Premium'})
            dtf1 = dtf.loc[:,
                   ['No', 'Policy', 'Endorsement No', 'Insured', 'From', 'To', 'Currency','Premium', 'RI Premium', 'RI Comm',
                    'RI Tax', 'Net RI Premium', 'File name']]
            return dtf1

        elif dtf.shape[1] == 18:
            dtf = dtf[[0, 1, 2, 5, 7,8, 12, 13, 14, 15, 'File name']]
            dtf = dtf.rename(
                columns={0: 'No', 1: 'Policy', 2: 'Insured', 5: 'From',7:'Premium', 8: 'Currency', 12: 'RI Premium', 13: 'RI Comm',
                         14: 'RI Tax', 15: 'Net RI Premium'})
            dtf1 = dtf.loc[:,
                   ['No', 'Policy', 'Insured', 'From','Premium', 'Currency', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium',
                    'File name']]
            dtf1.insert(2, 'Endorsement No', 'NA')
            dtf1.insert(5, 'To', '01/01/1900')
            return dtf1

        else:
            return dtf

#---------------------------------Get dataframe from claim files

def get_dataframe_claim(dtf):
    global line_or_text_method

    if line_or_text_method == 'text':
        if dtf.shape[1] == 15:
            dtf = dtf[[0, 1, 2, 4, 5, 6, 7, 9, 10, 11, 12, 'File name']]
            dtf1 = dtf.rename(
                columns={0: 'Policy', 1: 'Claim', 2: 'Insured', 4: 'From', 5: 'To', 6: 'Date of Loss', 7: 'Currency',
                         9: 'Total Loss Amount', 10: 'Survey Fee', 11: 'Total Amount', 12: 'RI Amount'})
            return dtf1

        elif dtf.shape[1] == 17:
            dtf = dtf[[0, 1, 2, 6, 7, 8, 10, 11, 12, 14, 'File name']]
            dtf1 = dtf.rename(
                columns={0: 'Policy', 1: 'Claim', 2: 'Insured', 6: 'From', 7: 'Date of Loss', 8: 'Currency',
                         10: 'Total Loss Amount', 11: 'Survey Fee', 12: 'Total Amount', 14: 'RI Amount'})
            dtf1.insert(4, 'To', '01/01/1900')
            return dtf1
        else:
            return dtf

    else:
        if dtf.shape[1] == 17:
            dtf = dtf[[0, 1, 2, 3, 5, 6, 7, 8, 10, 11, 12, 14, 'File name']]
            dtf1 = dtf.rename(
                columns={0: 'No', 1: 'Policy', 2: 'Claim', 3: 'Insured', 5: 'From', 6: 'To', 7: 'Date of Loss',
                         8: 'Currency', 10: 'Total Loss Amount', 11: 'Survey Fee', 12: 'Total Amount', 14: 'RI Amount'})
            return dtf1

        elif dtf.shape[1] == 19:
            dtf = dtf[[0, 1, 2, 3, 7, 8, 9, 11, 12, 13, 15, 'File name']]
            dtf1 = dtf.rename(
                columns={0: 'No', 1: 'Policy', 2: 'Claim', 3: 'Insured', 7: 'From', 8: 'Date of Loss', 9: 'Currency',
                         11: 'Total Loss Amount', 12: 'Survey Fee', 13: 'Total Amount', 15: 'RI Amount'})
            dtf1.insert(5, 'To', '01/01/1900')
            return dtf1

        else:
            return dtf

#---------------------------------------------Get dataframe from XOL files

def get_dataframe_xol(dtf):
    global line_or_text_method

    if line_or_text_method == 'text':
        if dtf.shape[1] == 18:
            dtf = dtf[[0,1,2,4,6,7,13,14,15,16,'File name']]
            dtf1 = dtf.rename(columns={0:'Policy',1:'Endorsement No',2:'Insured',4:'From',6:'Premium',7:'Currency',13:'RI Premium',14:'RI Comm',15:'RI Tax',16:'Net RI Premium'})
            dtf1.insert(4, 'To', '01/01/1900')
            return dtf1

        elif dtf.shape[1] == 19:
            dtf = dtf[[0,1,2,4,5,7,8,14,15,16,17,'File name']]
            dtf1 = dtf.rename(columns={0:'Policy',1:'Endorsement No',2:'Insured',4:'From',5:'To',7:'Premium',8:'Currency',14:'RI Premium',15:'RI Comm',16:'RI Tax',17:'Net RI Premium'})
            return dtf1
        else:
            return dtf

    else: #line
        if dtf.shape[1] == 20:
            dtf = dtf[[0,1,2,3,5,7,8,14,15,16,17,'File name']]
            dtf1 = dtf.rename(columns={0:'No',1:'Policy',2:'Endorsement No',3:'Insured',5:'From',7:'Premium',8:'Currency',14:'RI Premium',15:'RI Comm',16:'RI Tax',17:'Net RI Premium'})
            dtf1.insert(5, 'To', '01/01/1900')
            return dtf1

        elif dtf.shape[1] == 21:
            dtf = dtf[[0,1,2,3,5,6,8,9,15,16,17,18,'File name']]
            dtf1 = dtf.rename(columns={0:'No',1:'Policy',2:'Endorsement No',3:'Insured',5:'From',6:'To',8:'Premium',9:'Currency',15:'RI Premium',16:'RI Comm',17:'RI Tax',18:'Net RI Premium'})
            return dtf1

        else:
            return dtf
            

#------------------------------------------------------------

def button_select_folder_click():# open PDFfolder
    global folder_path
    global file_or_folder
    folder_path = fd.askdirectory()
    file_or_folder='folder'

#-----------------------------------------------------------

def button_saving_folder_click():
    global saving_folder_path
    saving_folder_path = fd.askdirectory()

#-----------------------------------------------------------

def button_run_click():
    global df
    global df_concat
    global df_concat_tmp
    global df_concat_1
    from tkinter import messagebox as msg
    if premium_or_claim =='':
        msg.showinfo("Information", "Select Premium or Claim ")
    elif line_or_text_method=='':
        msg.showinfo("Information", "Select a Method")
    elif file_or_folder=='':
        msg.showinfo("Information", "File/Folder has not been selected.Please retry")
    elif saving_folder_path == '':
        msg.showinfo("Information", "Saving folder has not been selected.Please retry")
    elif saving_file_name.get()=='':
        msg.showinfo("Information", "Enter resulted file name")
    else :
        if premium_or_claim == 'premium':
            if file_or_folder=='folder':
                read_folder()
                df_concat = pd.DataFrame(columns=premium_columns)
                for i in range(len(df)):
                    if df[i].shape[1] in [17, 19, 20,21,22]:
                        df[i] = get_dataframe_premium(df[i])
                        df_concat = pd.concat([df_concat, df[i]])
                    else:
                        df_concat = df_concat
            else: # file_or_folder=='file'
                read_file()
                df_concat = pd.DataFrame(columns=premium_columns)
                for i in range(len(df)):
                    if df[i].shape[1] in [17, 19, 20,21,22]:
                        df[i] = get_dataframe_premium(df[i])
                        df_concat = pd.concat([df_concat, df[i]])
                    else:
                        df_concat = df_concat
            if line_or_text_method=='line':
                df_concat_1 = df_concat.loc[(df_concat['No']!='No') & (df_concat['No'].notnull()) & (df_concat['No']!='')]
                df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']]=df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']].replace(' ', '', regex=True).replace('','0',regex=True).astype('float64')
            else:
                df_concat_1 = df_concat.loc[(df_concat['Policy']!='') & (df_concat['Policy'].notnull()) & (~df_concat['Policy'].str.contains('Contract')) & (~df_concat['Policy'].str.contains('Policy'))]
                df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']]=df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']].replace(' ', '', regex=True).replace('','0',regex=True).astype('float64')

        elif premium_or_claim == 'claim':# Claim
            if file_or_folder=='folder':
                read_folder()
                df_concat = pd.DataFrame(columns=claim_columns)
                for i in range(len(df)):
                    if df[i].shape[1] in [15, 17,19]:
                        df[i] = get_dataframe_claim(df[i])
                        df_concat = pd.concat([df_concat, df[i]])
                    else:
                        df_concat = df_concat
            else: # file_or_folder=='file'
                read_file()
                df_concat = pd.DataFrame(columns=claim_columns)
                for i in range(len(df)):
                    if df[i].shape[1] in [15, 17,19]:
                        df[i] = get_dataframe_claim(df[i])
                        df_concat = pd.concat([df_concat, df[i]])
                    else:
                        df_concat = df_concat

            if line_or_text_method=='line':
                df_concat_1 = df_concat.loc[(df_concat['No']!='No') & (df_concat['No'].notnull())& (df_concat['No']!='')]
                df_concat_1[['Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']]=df_concat_1[['Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']].replace(' ', '', regex=True).replace('','0',regex=True).astype('float64')
            else:
                df_concat_tmp = df_concat.loc[(df_concat['Claim']!='') & (df_concat['Claim'].notnull()) & (df_concat['Claim']!='Claim No.')]
                df_concat_tmp[['Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']]=df_concat_tmp[['Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']].replace(' ', '', regex=True)
                claim_no_temp_1=list(df_concat_tmp['Claim'])
                claim_no_temp_2=list(df_concat_tmp['Claim'])
                claim_no_final=list(df_concat_tmp['Claim'])
                policy_temp=list(df_concat_tmp['Policy'])
                for i in range(len(claim_no_temp_1)):
                    if policy_temp[i]=='': # concanate if claim no is splitted into 02 rows
                        claim_no_final[i-1]=claim_no_temp_1[i-1]+claim_no_temp_2[i]
                    else:
                        claim_no_final[i]=claim_no_final[i]
                df_concat_tmp=df_concat_tmp.reset_index(drop=True)
                df_concat_tmp.insert(1, 'New Claim No', claim_no_final )
                #df_concat_tmp['New Claim No']=claim_no_final  
                df_concat_1=df_concat_tmp.loc[(df_concat_tmp['Policy']!='') & (df_concat_tmp['Policy'].notnull())]
                df_concat_1[['Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']]=df_concat_1[['Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']].replace(' ', '', regex=True).replace('','0',regex=True).astype('float64')

        else: # xol
            if file_or_folder=='folder':
                read_folder()
                df_concat = pd.DataFrame(columns=premium_columns)
                for i in range(len(df)):
                    if df[i].shape[1] in [18,19,20,21]:
                        df[i] = get_dataframe_xol(df[i])
                        df_concat = pd.concat([df_concat, df[i]])
                    else:
                        df_concat = df_concat
            else: # file_or_folder=='file'
                read_file()
                df_concat = pd.DataFrame(columns=premium_columns)
                for i in range(len(df)):
                    if df[i].shape[1] in [18,19,20,21]:
                        df[i] = get_dataframe_xol(df[i])
                        df_concat = pd.concat([df_concat, df[i]])
                    else:
                        df_concat = df_concat
            if line_or_text_method=='line':
                df_concat_1 = df_concat.loc[(df_concat['No']!='No') & (df_concat['No'].notnull()) & (df_concat['No']!='')]
                df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']]=df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']].replace(' ', '', regex=True).replace('','0',regex=True).astype('float64')
            else: #text
                df_concat_tmp = df_concat.loc[(df_concat['Policy']!='') & (df_concat['Policy'].notnull()) & (~df_concat['Policy'].str.contains('Contract')) & (~df_concat['Policy'].str.contains('Policy'))]
                df_concat_tmp[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']]=df_concat_tmp[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']].replace(' ', '', regex=True) 
                policy_no_temp_1=list(df_concat_tmp['Policy'])
                policy_no_temp_2=list(df_concat_tmp['Policy'])
                policy_no_final=list(df_concat_tmp['Policy'])
                for i in range(len(policy_no_temp_1)):
                    if len(policy_no_temp_1[i])<8: # concanate if claim no is splitted into 02 rows
                        policy_no_final[i-1]=policy_no_temp_1[i-1]+policy_no_temp_2[i]
                    else:
                        policy_no_final[i]=policy_no_final[i]
                df_concat_tmp=df_concat_tmp.reset_index(drop=True)
                df_concat_tmp.insert(1, 'New Policy No', policy_no_final)
                #df_concat_tmp['New Claim No']=claim_no_final  
                df_concat_1=df_concat_tmp.loc[(df_concat_tmp['Policy']!='') & (df_concat_tmp['Policy'].notnull()) & (df_concat_tmp['Currency']!='')]
                df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']]=df_concat_1[['Premium', 'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']].replace(' ', '', regex=True).replace('','0',regex=True).astype('float64')

    df_concat_1['Policy type']=df_concat_1['Policy'].apply(get_policy_type)
    df_concat_1['LOB code']=df_concat_1['Policy type'].apply(get_lob)
    df_concat_1['LOB name']=df_concat_1['Policy type'].apply(get_lob_name)
    df_concat_1['Department']=df_concat_1['Policy'].apply(get_dept_name)
    fx_order_num=set_fx_order_num(df_concat_1)
    df_concat_1=df_concat_1.reset_index(drop=True)
    df_concat_1.insert(df_concat_1.shape[1],'fx_order_num', fx_order_num)
    df_concat_final=pd.merge(df_concat_1,fx_table,how='left',on='fx_order_num')

    #df_concat_1.to_csv(f"{saving_folder_path}/{saving_file_name.get()}.csv", encoding="utf-8-sig", index=False)

    writer = ExcelWriter(f"{saving_folder_path}/{saving_file_name.get()}.xlsx")

    df_concat_final.to_excel(writer, sheet_name='BANG_KE')


    fx_table.to_excel(writer, sheet_name='TY_GIA')

    writer.save()
    msg.showinfo("Information", "Completed successfully!")


def button_open_saved_file_click():
    subprocess.Popen([f"{saving_folder_path}/{saving_file_name.get()}.xlsx"], shell=True)
    #win.destroy()

def button_select_file_click():
    global file_path
    global file_or_folder
    file_path=fd.askopenfilename()
    file_or_folder='file'


def rad_select_premium_click():
    global premium_or_claim
    premium_or_claim = rad_premium_claim_var.get() # return 'premium'


def rad_select_claim_click():
    global premium_or_claim
    premium_or_claim = rad_premium_claim_var.get() # return 'claim'


def rad_select_xol_click():
    global premium_or_claim
    premium_or_claim = rad_premium_claim_var.get() # return 'xol-pre'


def rad_line_method_click():
    global line_or_text_method
    global premium_columns
    global claim_columns
    line_or_text_method = rad_line_text_var.get() # return 'line'
    premium_columns = ['No', 'Policy', 'Endorsement No', 'Insured', 'From', 'To', 'Premium', 'Currency',
                       'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']
    claim_columns = ['No', 'Policy', 'Claim', 'Insured', 'From', 'To', 'Date of Loss', 'Currency',
                     'Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']

def rad_text_method_click():
    global line_or_text_method
    global premium_columns
    global claim_columns
    line_or_text_method = rad_line_text_var.get() #return 'text'
    premium_columns = ['Policy', 'Endorsement No', 'Insured', 'From', 'To', 'Premium', 'Currency',
                       'RI Premium', 'RI Comm', 'RI Tax', 'Net RI Premium']
    claim_columns = ['Policy', 'Claim', 'Insured', 'From', 'To', 'Date of Loss', 'Currency',
                     'Total Loss Amount', 'Survey Fee', 'Total Amount', 'RI Amount']


rad_premium_claim_var= tk.StringVar()
rad_line_text_var = tk.StringVar()


rad_select_premium = ttk.Radiobutton(win, text='Premium', width=30, value='premium', variable=rad_premium_claim_var, command=rad_select_premium_click)
rad_select_premium.grid(column=0, row=0)

rad_select_claim = ttk.Radiobutton(win, text='Claim', width=30, value='claim', variable=rad_premium_claim_var, command=rad_select_claim_click)
rad_select_claim.grid(column=1, row=0)

rad_select_xol = ttk.Radiobutton(win, text='XOL Premium', width=30, value='xol-pre', variable=rad_premium_claim_var, command=rad_select_xol_click)
rad_select_xol.grid(column=2, row=0)

rad_line_method = ttk.Radiobutton(win, text='Lines method', width=30, value='line', variable=rad_line_text_var, command=rad_line_method_click)
rad_line_method.grid(column=0, row=1)

rad_text_method = ttk.Radiobutton(win, text='Text method', width=30, value='text', variable=rad_line_text_var, command=rad_text_method_click)
rad_text_method.grid(column=1, row=1)

label_select_folder = ttk.Label(win, text="Select PDF Folder", width=30).grid(column=0, row=4)
button_select_folder = ttk.Button(win, text="PDF Folder",width=16, command=button_select_folder_click)
button_select_folder.grid(column=1, row=4)

label_select_saving_folder = ttk.Label(win, text="Select saving folder", width=30).grid(column=0, row=6)
button_saving_folder = ttk.Button(win, text="Saving folder",width=16, command=button_saving_folder_click)
button_saving_folder.grid(column=1, row=6)

label_saving_file_name = ttk.Label(win, text="Enter saving file name", width=30).grid(column=0, row=8)
saving_file_name = ttk.Entry(win, width=16)
saving_file_name.grid(column=1, row=8)

label_run = ttk.Label(win, text="Process", width=30).grid(column=0, row=12)
button_run = ttk.Button(win, text="Run",width=16, command=button_run_click)
button_run.grid(column=1, row=12)

label_open_saved_file = ttk.Label(win, text="Open saved file", width=30).grid(column=0, row=16)
button_open_saved_file = ttk.Button(win, text="Open saved file",width=16, command=button_open_saved_file_click)
button_open_saved_file.grid(column=1, row=16)

label_select_file = ttk.Label(win, text="Open file", width=30).grid(column=0, row=3)
button_select_file = ttk.Button(win, text="Select file",width=16, command=button_select_file_click)
button_select_file.grid(column=1, row=3)

win.mainloop()
