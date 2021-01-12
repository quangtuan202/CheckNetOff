import pandas as pd 
from fuzzywuzzy import fuzz 
from fuzzywuzzy import process
def checker(wrong_options,correct_options):
    names_array=[]
    ratio_array=[]    
    for wrong_option in wrong_options:
        if wrong_option in correct_options:
           names_array.append(wrong_option)
           ratio_array.append('100')
        else:   
            x=process.extractOne(wrong_option,correct_options,scorer=fuzz.token_set_ratio)
            names_array.append(x[0])
            ratio_array.append(x[1])
    return names_array,ratio_array
df1 = pd.read_excel('D:/05AMSIGv2.xlsx')
df2=pd.read_excel('D:/MSIG.xlsx')
df1List=df1['KH&Phi'].tolist()
df2List=df2['KH&Phi'].tolist()
name_match,ratio_match=checker(df1List,df2List)
df3 = pd.DataFrame()
df3['old_names']=pd.Series(df1List)
df3['correct_names']=pd.Series(name_match)
df3['correct_ratio']=pd.Series(ratio_match)
df4=df3.merge(df1,left_on='old_names',right_on='KH&Phi',how='left')
df4.to_excel('D:/matched_names.xlsx', engine='xlsxwriter')