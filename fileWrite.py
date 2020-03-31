import pandas as pd
from openpyxl import load_workbook

def writeToFile(path,obj,trans=1):
    print('Writing to file : ',path)
    writer = pd.ExcelWriter(path,mode = "a") 
    # print('writer')
    for key in obj:
        if (trans):
            k = map(list,zip(*obj[key]))
        else:
            k = obj[key]
        write = pd.DataFrame(k)
        write.to_excel(writer,key,header=False,index=False)
        # print('Writing : ',key)
    writer.save()
    # print('save')
