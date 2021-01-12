import pandas as pd
filename=r"D:\OneDrive\OneDrive - khoavanhoc.edu.vn\UIC\SQL_data\2020\202010\BCTH07A_01-10 2020 12112020 1010.xlsx"
filename.replace('\\','/')
df=pd.read_excel(filename)
df['Môi giới phí BHG']=df['Môi giới phí BHG'].fillna(0)
df['Hoa hồng nhận tái bảo hiểm']=df['Hoa hồng nhận tái bảo hiểm'].fillna(0)
df['Hoa hồng nhận tái bảo hiểm']=df['Hoa hồng nhận tái bảo hiểm'].fillna(0)
df['Hoa hồng đại lý bảo hiểm']=df['Hoa hồng đại lý bảo hiểm'].fillna(0)
dfSelect=df.loc[(df['Phí bảo hiểm \n(VND)']!=0),['Số đơn/Endor','Mã nghiệp vụ','Phí bảo hiểm \n(VND)','Hoa hồng đại lý bảo hiểm','Môi giới phí BHG','Hoa hồng nhận tái bảo hiểm']]
dfSelect['Tong hoa hong']=df['Môi giới phí BHG']+df['Hoa hồng đại lý bảo hiểm']+df['Hoa hồng nhận tái bảo hiểm']
dfSelect=dfSelect[dfSelect['Tong hoa hong']!=0]
dfSelect['Ty le Hoa hong cua don']=dfSelect['Tong hoa hong']/df['Phí bảo hiểm \n(VND)']
dfSelect['Ty le Hoa hong']=dfSelect.groupby(['Mã nghiệp vụ'])['Tong hoa hong'].transform('sum')/dfSelect.groupby(['Mã nghiệp vụ'])['Phí bảo hiểm \n(VND)'].transform('sum')
dfSelect['Mean']=dfSelect.groupby(['Mã nghiệp vụ'])['Ty le Hoa hong'].transform('mean')
dfSelect['Median']=dfSelect.groupby(['Mã nghiệp vụ'])['Ty le Hoa hong'].transform('median')
dfSelect['Std']=dfSelect.groupby(['Mã nghiệp vụ'])['Ty le Hoa hong cua don'].transform('std',ddof=0)
dfSelect['Z-Score']=((dfSelect['Ty le Hoa hong cua don']-dfSelect['Mean'])/dfSelect['Std']).abs()
dfSelect.to_csv("D:/checkHHgoc.csv",encoding='utf-8-sig')
save_file="D:/checkHHgoc.csv"
dfSelect.to_csv(save_file)
subprocess.Popen([save_file],shell=True)