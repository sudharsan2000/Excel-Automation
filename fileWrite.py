import pandas as pd
from openpyxl import load_workbook

def writeToFile(path,obj,trans=1):

    writer = pd.ExcelWriter(path, engine='openpyxl') 

    for key in obj:
        if (trans):
            k = map(list,zip(*obj[key]))
        else:
            k = obj[key]
        write = pd.DataFrame(k)
        write.to_excel(writer,key,header=False,index=False)
    writer.save()