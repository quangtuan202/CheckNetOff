import xml.etree.ElementTree as ET
import pandas as pd
file='D:/hoadon/hoadon (10).inv'
tree=ET.parse(file)
root = tree.getroot()
lst=[]
for _property in root.findall('property'): 
    kind = _property.find('kind').text
    value = _property.find('value').text
    content=kind+' '+value
    lst.append(content)
finalList=[]
for line in lst:
    if line.startswith('So don'):
        finalList.append(line)
df=pd.DataFrame([finalList],columns=['So don cua Leader','So don UIC'])
df['So don cua Leader']=df['a'].str.split('Leader',expand=True)[1]
df['So don cua UIC']=df['b'].str.split('BH',expand=True)[1]
df=df.loc[:,['So don cua Leader','So don cua UIC']]
