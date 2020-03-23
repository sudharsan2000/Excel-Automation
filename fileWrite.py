import pandas as pd
from openpyxl import load_workbook

def writeToFile(path,obj):

    # book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine='openpyxl') 
    # writer.book = book
    for key in obj:
        k = map(list,zip(*obj[key]))
        write = pd.DataFrame(k)
        write.to_excel(writer,key)
    writer.save()