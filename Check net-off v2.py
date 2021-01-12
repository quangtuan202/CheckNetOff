# optimization for speed
# calculate first one-third, last one third to reduce workload
import pandas as pd 
import numpy as np
from itertools import combinations 
import xlrd
import pyxlsb
import datetime as dt

# Clean policy with dot and Hyphen at the end
def clean_policy(x):
    if x[-1] in ('.','-'):
        return x[:-1]
    else:
        return x
# This function is applicable to Marine policy
def right_find(policyCol, cpcAmountCol, accAmountCol):
    condition = [cpcAmountCol!=0 and accAmountCol==0 and policyCol.rfind('.')>policyCol.rfind('-'),
                cpcAmountCol!=0 and accAmountCol==0 and policyCol.rfind('.')<policyCol.rfind('-')]

    result=[policyCol[0:policyCol.rfind('.')],
            policyCol[0:policyCol.rfind('-')]]
    return np.select(condition,result,policyCol)

# Function to check if there is negative item
def check_negative(iterable):
    lst=[x for x in iterable if x<0]
    if len(lst)==1 and sum(iterable)<-1000: # If sum of a combination < limit ==> no net-off amount
        return False
    else:
        return True

# Function to create an iterable
def create_iterable(length):
    x=[i for i in range(length,1,-1)]
    y=[i for i in range(2,length+1)]
    z=list(zip(x,y))
    return [x for y in z for x in y][0:len(z)]

# d = {k:v for v in a for k in v} create dict for mapping with index
# Function to check net-off amount for Non Marine
def netOffCheckNonMarine(x,df,col,totalLimit,combinationLimit):
    #col : columns to select
    #threshold : amount limit
    #global combinationList
    global global_index_list
    separateGroup=df.loc[x.index,col]
    separateGroup=separateGroup.sort_values(col)
    dict=separateGroup.to_dict()
    index_list=[]
    print('separateGroup:Checking...')
    print('Number of rows: ',len(separateGroup))
    print(separateGroup.head(1))
    print('\n')
    
    for i in create_iterable(len(dict[col[0]])):
        print('len of dict:' ,len(dict[col[0]]))
        print('Combination:', i)
        if len(dict[col[0]])>1 and check_negative(list(dict[col[1]].values()))==True:
            if (i>len(dict[col[0]])) or comb(len(dict[col[0]].keys()),i)>combinationLimit:
                continue
            index_com=combinations(dict[col[0]].keys(),i)
            print('combination of:',len(dict[col[0]].keys()),i)
            print(dict[col[0]].keys())
            amount_com=combinations(dict[col[1]].values(),i)
            for j,k in zip(amount_com,index_com):
                # if first item of each combination > 0=> no net-off ( df is sorted ascending)
                if j[0]>0:
                    break
                #if sum of an array=0 and each array does not contain an item of another array
                if np.absolute(np.sum(np.array(j)))<=totalLimit and len([x for x in k if x in [ y for z in index_list for y in z]])==0:
                    index_list.append(k)
                    global_index_list.append(k)
                    print('Sum=0 found:', k)
                    #combinationList.append(k) # add the combination to list
                    for l in k:
                        del dict[col[0]][l]
                        del dict[col[1]][l]

    index_list=[x for y in index_list for x in y]
    index_list.sort()
    index_list2=["Net-off" if i in index_list  else "" for i in x.index]
    print('Done!')
    return index_list2


def netOffCheckMarine(x,df,col,totalLimit,rowLimit):
    #col : columns to select
    #totalLimit : amount limit
    #rowLimit : total number of row to make combinations, if rowLimit> 30, if is very hard to calculate using PC
    global combinationList
    separateGroup=df.loc[x.index,col]
    separateGroup=separateGroup.sort_values(col)
    dict=separateGroup.to_dict()
    index_list=[]
    print('separateGroup:Checking...')
    print(separateGroup.head(1))
    print('\n')
    for i in create_iterable(len(dict[col[0]])):
        if len(dict[col[0]])>1 and len(dict[col[0]])<rowLimit and check_negative(list(dict[col[1]].values()))==True:
            if i>len(dict[col[0]]):
                continue
            index_com=combinations(dict[col[0]].keys(),i)
            amount_com=combinations(dict[col[1]].values(),i)
            for j,k in zip(amount_com,index_com):
                  # if first item of each combination > 0=> no net-off ( df is sorted ascending)
                if j[0]>0:
                    break
                print('combination: ',k)
                print('Sum:',sum(j))
                #if sum of an array=0 and each array does not contain an item of another array
                if np.absolute(np.sum(np.array(j)))<=totalLimit and len([x for x in k if x in [ y for z in index_list for y in z]])==0:
                    index_list.append(k)
                    combinationList.append(k) # add the combination to list
                    for l in k:
                        del dict[col[0]][l]
                        del dict[col[1]][l]

    index_list=[x for y in index_list for x in y]
    index_list.sort()
    index_list2=["Net-off" if i in index_list  else "" for i in x.index]
    print('Done!')
    return index_list2

#Set threshold for Marine to filter marine policy with diff amount to 3% or bigger
marineThreshold=0.03 

#Create an empty list to store index combinations
combinationList=[]


# Read file ,dfMarinePolicy contains Marine Policy needs to be checked
dfCheckSQL=pd.read_excel('D:/Check_GWP.xlsx',usecols=['Nhan tai/Truc tiep','Số đơn','Số đơn đến mã phòng','Mã nghiệp vụ ','Phòng','Phí bảo hiểm NT CPC','Phí bảo hiểm NT ACC','Chênh lệch NT','Phí bảo hiểm VND CPC','Phí bảo hiểm VND ACC','Chênh lệch VND','Final Note'])

# Clean dfCheckSQL
dfCheckSQL['Số đơn']=dfCheckSQL['Số đơn'].apply(clean_policy) # clean white space
dfCheckSQL['Số đơn đến mã phòng']=dfCheckSQL['Số đơn đến mã phòng'].apply(clean_policy) # clean white space
#dfCheckSQL['Số đơn']=np.where(dfCheckSQL['Số đơn'].str.contains('UPDATE'),dfCheckSQL['Số đơn'].str.slice(start=7),dfCheckSQL['Số đơn'])
#dfCheckSQL['Số đơn đến mã phòng']=np.where(dfCheckSQL['Số đơn đến mã phòng'].str.contains('UPDATE'),dfCheckSQL['Số đơn đến mã phòng'].str.slice(start=7),dfCheckSQL['Số đơn đến mã phòng'])
#MarinePolicy=pd.read_excel('D:/Check_GWP.xlsx',sheet_name='MarinePolicy').values()

# Define col for 02 groups
colNonMarine=['Số đơn đến mã phòng','Chênh lệch VND']
colMarine=['PolicyForChecking','Chênh lệch VND']

# Create df for Non Marine
dfNonMarine=dfCheckSQL.loc[(~dfCheckSQL['Mã nghiệp vụ '].isin(['MCA','ACA','MCE','MCI','ICA','COT',' MC']))&(~dfCheckSQL['Final Note'].str.contains('OK'))]
dfNonMarine.sort_values(colNonMarine,inplace=True)

dfNonMarine['Số đơn']=dfNonMarine['Số đơn'].str.strip() # clean white space
dfNonMarine['Số đơn đến mã phòng']=dfNonMarine['Số đơn đến mã phòng'].str.strip() # clean white space
dfNonMarine['Số đơn']=np.where(dfNonMarine['Số đơn'].str.contains('UPDATE'),dfNonMarine['Số đơn'].str.slice(start=7),dfNonMarine['Số đơn'])
dfNonMarine['Số đơn đến mã phòng']=np.where(dfNonMarine['Số đơn đến mã phòng'].str.contains('UPDATE'),dfNonMarine['Số đơn đến mã phòng'].str.slice(start=7),dfNonMarine['Số đơn đến mã phòng'])


# Create df for Marine
dfMarine=dfCheckSQL.loc[(dfCheckSQL['Mã nghiệp vụ '].isin(['MCA','ACA','MCE','MCI','ICA','COT',' MC']))&(~dfCheckSQL['Final Note'].str.contains('OK'))]


# Create df for MOT
#dfMOT=dfCheckSQL.loc[(dfCheckSQL['Mã nghiệp vụ '].isin(['MOT']))&(dfCheckSQL['Final Note']).isin(['Check','Acc chưa ghi nhận','CLTG-Cargo'])]
#dfMOT['Policy']=dfMOT.apply(lambda x: right_find(x['Số đơn'],x['Phí bảo hiểm VND CPC'],x['Phí bảo hiểm VND ACC']),axis=1).astype(str)
#dfMOT['Check']=dfMOT.groupby([colNonMarine[0]],sort=False)[colNonMarine[1]].transform(netOffCheckNonMarine,df=dfMOT,col=colNonMarine,totalLimit=5000)
#dfMOT['Check2']=dfMOT.groupby(['Policy'],sort=False)[colNonMarine[1]].transform(netOffCheckNonMarine,df=dfMOT,col=colNonMarine,totalLimit=5000)

# Create df group by for Marine
dfMarineGroupBy=dfMarine.groupby(['Số đơn đến mã phòng'],as_index=False,sort=False)['Phí bảo hiểm VND CPC','Phí bảo hiểm VND ACC'].sum()
dfMarineGroupBy['PercentageDiff']=(2*(dfMarineGroupBy['Phí bảo hiểm VND CPC']-dfMarineGroupBy['Phí bảo hiểm VND ACC']+1)/(dfMarineGroupBy['Phí bảo hiểm VND CPC']+dfMarineGroupBy['Phí bảo hiểm VND ACC'])).abs()

marineList=dfMarineGroupBy.loc[dfMarineGroupBy['PercentageDiff']>marineThreshold,['Số đơn đến mã phòng']]
marineList=marineList['Số đơn đến mã phòng'].values

# Filter dfMarine 
dfMarineForChecking=dfMarine.loc[dfMarine['Số đơn đến mã phòng'].isin(marineList)]
dfMarineForChecking['Số đơn']=dfMarineForChecking['Số đơn'].str.strip() # clean white space
dfMarineForChecking['Số đơn đến mã phòng']=dfMarineForChecking['Số đơn đến mã phòng'].str.strip() # clean white space
dfMarineForChecking['Số đơn']=np.where(dfMarineForChecking['Số đơn'].str.contains('UPDATE'),dfMarineForChecking['Số đơn'].str.slice(start=7),dfMarineForChecking['Số đơn'])
dfMarineForChecking['Số đơn đến mã phòng']=np.where(dfMarineForChecking['Số đơn đến mã phòng'].str.contains('UPDATE'),dfMarineForChecking['Số đơn đến mã phòng'].str.slice(start=7),dfMarineForChecking['Số đơn đến mã phòng'])
dfMarineNOTForChecking=dfMarine.loc[~dfMarine['Số đơn đến mã phòng'].isin(marineList)]
dfMarineNOTForChecking['Check']=dfMarineNOTForChecking['Final Note']
dfMarineForChecking['PolicyForChecking']=dfMarineForChecking.apply(lambda x: right_find(x['Số đơn'],x['Phí bảo hiểm VND CPC'],x['Phí bảo hiểm VND ACC']),axis=1).astype(str)
dfMarineForChecking.sort_values(colMarine,inplace=True)
#dfMarineForChecking=dfMarineForChecking.sort_values(colMarine)

# Create df for the rest
#dfOk=dfCheckSQL.loc[(~dfCheckSQL['Final Note'].isin(['Check','Acc chưa ghi nhận','CLTG-Cargo']))]
#dfOk['Check']=dfOk['Final Note']

# Create column for Marine policy

dfNonMarine['Check']=dfNonMarine.groupby([colNonMarine[0]],sort=False)[colNonMarine[1]].transform(netOffCheckNonMarine,df=dfNonMarine,col=colNonMarine,totalLimit=10000)

dfMarineForChecking['Check']=dfMarineForChecking.groupby([colMarine[0]],sort=False)[colMarine[1]].transform(netOffCheckMarine,df=dfMarineForChecking,col=colMarine,totalLimit=20000,rowLimit=101)

#dfCheckFinal=pd.concat([dfNonMarine,dfMarineForChecking,dfMarineNOTForChecking,dfOk],axis=0)
dfCheckFinal=pd.concat([dfNonMarine,dfMarineForChecking,dfMarineNOTForChecking],axis=0)
dfCheckFinal['Note']=np.where(dfCheckFinal['Check']=='Net-off',dfCheckFinal['Check'],dfCheckFinal['Final Note'])
dfCheckFinal=dfCheckFinal.drop(columns=['Final Note','Check'])
dfCheckLastMonth=pd.read_excel("D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/DOI CHIEU ACC-CPC/2020/01-10/GWP 01-10 2020.xlsb",usecols=['Số đơn','Final Note'],skiprows=1,engine='pyxlsb')
dfCheckLastMonth=dfCheckLastMonth.rename(columns={'Final Note':'Note thang truoc'})
dfCheckFinal=dfCheckFinal.merge(dfCheckLastMonth, on='Số đơn',how='left')
time=str(dt.datetime.now()).replace(':','.')
dfCheckFinal.to_csv(f'D:/dfCheckFinal{time}.csv',encoding='utf-8-sig')
import subprocess
subprocess.Popen([f'D:/dfCheckFinal{time}.csv'],shell=True)
print('Complete!!!')

