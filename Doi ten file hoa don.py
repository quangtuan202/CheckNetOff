import xml.etree.ElementTree as ET
import pandas as pd
import os
folder='D:/hoadon'

# Real xml file
for filename in os.listdir(folder):
    tree=ET.parse(f'{folder}/{filename}')
    root = tree.getroot()
    for invoice in root.findall('Content'):
        no = invoice.find('InvoiceNo').text
        try:
            os.rename(f'{folder}/{filename}',f'{folder}/{no}.inv')
        except Exception:
            pass