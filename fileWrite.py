import pandas as pd
from openpyxl import load_workbook

def writeToFile(path,original,obj,trans=1):
    print('Writing to file : ',path)
    writer = pd.ExcelWriter(path,mode = "w") 
    for key in original:
        k = original[key]
        #write = pd.DataFrame(k)
        k.to_excel(writer,key,header=False,index=False) 
    for key in obj:
        if (trans):
            k = map(list,zip(*obj[key]))
        else:
            k = obj[key]
        write = pd.DataFrame(k)
        write.to_excel(writer,key,header=False,index=False)
    
        
    writer.save()
    # print('save')
