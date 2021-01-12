import pandas as pd
b2020="D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/SQL_data/2020/202011/bc_th_05b 01-11 2020.XLS"
b2019="D:/OneDrive/OneDrive - khoavanhoc.edu.vn/UIC/SQL_data/2019.12M/20200311/bc_th_05b.xlsx"
c="D:/UIC - Số liệu bồi thường update T6789 năm 2020.xlsx"
df2019=pd.read_excel(b2019,skiprows=1,header=None)
df2020=pd.read_excel(b2020,skiprows=8,header=None)
dfc=pd.read_excel(c)
df2019=df2019.iloc[:,[4,5,7,10,12,14,17,25,26,27,28]]
df2020=df2020.iloc[:,[4,5,7,10,12,14,17,25,26,27,28]]
col=['Số đơn/ Số Endor','Số hồ sơ bồi thường','Tên Khách hàng','Loại tiền nguyên tệ','Bồi thường(VND)','Phí giám định','Account month','Thu Bồi thườn nhượng tái bảo hiểm (cho các cty NN)','Thu Bồi thường từ nhượng tái bảo hiểm (cho các cty VN)','Thu giảm bồi thường','Bồi thường thuộc trách nhiệm giữ lại']
df2019.columns=col
df2020.columns=col
dfClaim=pd.concat([df2019,df2020],axis=0)
dfc['PolicyNo']=dfc['PolicyNo'].astype(str)
lst=dfc['PolicyNo'].values

def matcher(x):
    for i in lst:
        if i.lower() in x.lower():
            return i
    else:
        return 'no'

dfClaim['Số hồ sơ bồi thường']=dfClaim['Số hồ sơ bồi thường'].astype(str)
dfClaim['check']=dfClaim['Số hồ sơ bồi thường'].apply(matcher)
dfClaim.to_csv('D:/CheckClaim.csv',encoding='utf-8-sig')