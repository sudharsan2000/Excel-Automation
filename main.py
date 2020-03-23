import pandas as pd
COS={'CO1':[],'CO2':[],'CO3':[]}
COS_count={'CO1':0,'CO2':0,'CO3':0}
try:
    excel = pd.ExcelFile('CO-PO MOS U15AET502 -odd 2018-19-Micro.xlsx')
    INTS = [pd.read_excel(excel,'INT1'), pd.read_excel(excel,'INT2') ]
except:
    print('Error reading file')
    exit(1)

NAMES = INTS[0].iloc[7:,2]
ROLL_NO = INTS[0].iloc[7:,1]
# INT1_CONTENTS = {}
# INT1_CONTENTS['NAMES'] = INT1.iloc[7:,2]
# INT1_CONTENTS['T1-Q' + str(i)] = INT1.iloc[3:,2+i]
# print(INT1)
# print(INT1_CONTENTS['NAMES'])
def readInternals(INTS):
    iter = 1
    for INT in INTS:
        print('Finding COs from INT',iter)
        iter+=1
        for i in range(1,1000):
            try:
                COS[str(INT.iat[5,2+i])].append(INT.iloc[3:,2+i])
                # print(INT.iat[5,2+i])
                COS_count[INT.iat[5,2+i]] +=1
            except:
                break
    print('Done!')
    print('Number of CO1 questions found : ',COS_count['CO1'])
    print('Number of CO2 questions found : ',COS_count['CO2'])
    print('Number of CO3 questions found : ',COS_count['CO3'])

# print(COS)
def main():
    readInternals(INTS)


if( __name__ == "__main__"):
    main()
