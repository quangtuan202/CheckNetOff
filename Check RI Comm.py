import pandas as pd
import subprocess
filename=r"D:\OneDrive\OneDrive - khoavanhoc.edu.vn\UIC\SQL_data\2020\202010\BCTH05A_01-10 2020 10112020.xlsx"
filename.replace('\\','/')
df=pd.read_excel(filename)
df['So don bo EN']=df['Số đơn/ Số Endor'].str.split('EN',n=1,expand=True)[0]
df['Tong phi nhuong']=(df['Phí nhượng TBH FAC']+df['Phí nhượng TBH QS']+ df['Phí nhượng TBH Surplus ']).fillna(0)
df['Tong HH nhuong']=df['Hoa hồng nhượng TBH FAC']+ df['Hoa hồng nhượng TBH QS'] +df['Hoa hồng nhượng TBH Surplus '].fillna(0)
df['Tong HH nhuong don bo EN']=df.groupby(['So don bo EN'],as_index=False)['Tong HH nhuong'].transform(sum)
df['Tong phi nhuong bo EN']=df.groupby(['So don bo EN'],as_index=False)['Tong phi nhuong'].transform(sum)
df['Ty le HH nhuong cua don']=df['Tong HH nhuong']/df['Tong phi nhuong']
df['Ty le HH nhuong cua don bo EN']=df['Tong HH nhuong don bo EN']/df['Tong phi nhuong bo EN']
df['phi nhuong theo nhom']=df.groupby(['Mã nghiệp vụ'],as_index=False)['Tong phi nhuong'].transform(sum)
df['HH nhuong theo nhom']=df.groupby(['Mã nghiệp vụ'])['Tong HH nhuong'].transform(sum)
df['Ty le HH nhuong theo nhom']=df['HH nhuong theo nhom']/df['phi nhuong theo nhom']
condition=((df['Tong phi nhuong']!=0) & (df['Tong phi nhuong'].notna()) & (df['Tong HH nhuong']!=0) & (df['Tong HH nhuong'].notna()))
dfSelect=df.loc[condition,['Số đơn/ Số Endor','Mã nghiệp vụ','Tong phi nhuong','Tong HH nhuong','Ty le HH nhuong cua don','Tong phi nhuong bo EN','Tong HH nhuong don bo EN','Ty le HH nhuong cua don bo EN','phi nhuong theo nhom','HH nhuong theo nhom','Ty le HH nhuong theo nhom']]
# calculate population standard deviation ( ddof=0), sample standard deviation (ddof=1)
dfSelect['STD1']=dfSelect.groupby('Mã nghiệp vụ')['Ty le HH nhuong cua don'].transform('std',ddof=0)
dfSelect['Mean1']=dfSelect.groupby('Mã nghiệp vụ')['Ty le HH nhuong cua don'].transform('mean')
dfSelect['Z-Score']=((dfSelect['Ty le HH nhuong cua don']-dfSelect['Mean1'])/dfSelect['STD1']).abs()
dfSelect['STD2']=dfSelect.groupby('Mã nghiệp vụ')['Ty le HH nhuong cua don bo EN'].transform('std',ddof=0)
dfSelect['Mean2']=dfSelect.groupby('Mã nghiệp vụ')['Ty le HH nhuong cua don bo EN'].transform('mean')
dfSelect['Z-Score2']=((dfSelect['Ty le HH nhuong cua don bo EN']-dfSelect['Mean2'])/dfSelect['STD2']).abs()
save_file="D:/checktlhh.csv"
dfSelect.to_csv(save_file)
subprocess.Popen([save_file],shell=True)