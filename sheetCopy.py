from openpyxl import load_workbook
import pandas as pd
wb = load_workbook(filename = 'CO-PO MOS U15AET501 -odd 2018-19-Micro.xlsx')
sheet_names = wb.get_sheet_names()
name = sheet_names[0]
sheet_ranges = wb[name]
df = pd.DataFrame(sheet_ranges.values)
print(df)