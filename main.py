import pandas as pd


try:
    excel = pd.ExcelFile('CO-PO MOS U15AET502 -odd 2018-19-Micro.xlsx')
    INT1 = pd.read_excel(excel,'INT1') 
    INT2 = pd.read_excel(excel,'INT2')
except:
    print('Error reading file')
    exit(1)

print(INT1)