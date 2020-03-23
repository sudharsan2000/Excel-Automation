import pandas as pd
COS={'CO1':[],'CO2':[],'CO3':[]}
COS_count={'CO1':0,'CO2':0,'CO3':0}
try:
    excel = pd.ExcelFile('CO-PO MOS U15AET502 -odd 2018-19-Micro.xlsx')
    INT1 = pd.read_excel(excel,'INT2') 
    INT2 = pd.read_excel(excel,'INT2')
except:
    print('Error reading file')
    exit(1)
# INT1 = INT1.iloc[3:,:]
INT1_CONTENTS = {}
INT1_CONTENTS['NAMES'] = INT1.iloc[7:,2]
# print(INT1)
# print(INT1_CONTENTS['NAMES'])

for i in range(1,20):
    try:
        INT1_CONTENTS['T1-Q' + str(i)] = INT1.iloc[3:,2+i]
        COS[str(INT1.iat[5,2+i])].append(INT1.iloc[3:,2+i])
        print(INT1.iat[5,2+i])
        COS_count[INT1.iat[5,2+i]] +=1
    except:
        break

# for i in range(1,20):
#     print(INT1_CONTENTS['T1-Q' + str(i)])

# print(INT1.iat[5,2+1])
print(COS_count['CO1'])
print(COS_count['CO2'])
print(COS_count['CO3'])


