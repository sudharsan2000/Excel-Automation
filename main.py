import pandas as pd
import numpy as np
from fileWrite import * 

COS={'CO1':[],'CO2':[],'CO3':[],'CO4':[],'CO5':[],'CO6':[]}
COS_count={'CO1':0,'CO2':0,'CO3':0,'CO4':0,'CO5':0,'CO6':0}
MARKS = []
try:
    excel = pd.ExcelFile('CO-PO MOS U15AET502 -odd 2018-19-Micro.xlsx')
    INTS = [pd.read_excel(excel,'INT1'), pd.read_excel(excel,'INT2') ]
except:
    print('Error reading file')
    exit(1)

NAMES = INTS[0].iloc[3:,2]
ROLL_NO = INTS[0].iloc[7:,1]

def readInternals():
    iter = 1
    COS['CO1'].append(INTS[0].iloc[3:,1])
    COS['CO1'].append(INTS[0].iloc[3:,2])

    COS['CO2'].append(INTS[0].iloc[3:,1])
    COS['CO2'].append(INTS[0].iloc[3:,2])

    COS['CO3'].append(INTS[0].iloc[3:,1])
    COS['CO3'].append(INTS[0].iloc[3:,2])

    COS['CO4'].append(INTS[0].iloc[3:,1])
    COS['CO4'].append(INTS[0].iloc[3:,2])

    COS['CO5'].append(INTS[0].iloc[3:,1])
    COS['CO5'].append(INTS[0].iloc[3:,2])

    COS['CO6'].append(INTS[0].iloc[3:,1])
    COS['CO6'].append(INTS[0].iloc[3:,2])

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
    print('Number of CO3 questions found : ',COS_count['CO4'])
    print('Number of CO3 questions found : ',COS_count['CO5'])
    print('Number of CO3 questions found : ',COS_count['CO6'])


# print(COS)
MARKS = []
MARKS_PER_QUESTION = []
QUESTION_ATTEMPTED = []
def main():
    global MARKS
    readInternals()
    # print(COS['CO1'])
    writeToFile(r"./Course Outcomes.xlsx",COS)
    iter = 1

    for flag in COS:
        MARKS.append((pd.DataFrame(COS[flag]).to_numpy()[2:COS_count[flag] + 3,4:]).transpose() )
        MARKS_PER_QUESTION.append((pd.DataFrame(COS[flag]).to_numpy()[2:COS_count[flag] + 3,0]).transpose() )

    print(MARKS_PER_QUESTION)
    print(np.nansum(MARKS[0],axis=1))
    k = np.array(MARKS[0],dtype=float)
    print(k.dtype)
    print(np.count_nonzero(~np.isnan(k[0])))
if( __name__ == "__main__"):
    main()
