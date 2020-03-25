import pandas as pd
import numpy as np
from fileWrite import * 
import scipy.sparse as sparse

COS={'CO1':[],'CO2':[],'CO3':[],'CO4':[],'CO5':[],'CO6':[]}
COS_count={'CO1':0,'CO2':0,'CO3':0,'CO4':0,'CO5':0,'CO6':0}
GRADE_RANGES=[60,50,40]
MARKS = []
try:
    excel = pd.ExcelFile('CO-PO MOS U15AET502 -odd 2018-19-Micro.xlsx')
    INTS = [pd.read_excel(excel,'INT1'), pd.read_excel(excel,'INT2') ]
except:
    print('Error reading file')
    exit(1)

NAMES = INTS[0].iloc[3:,2]
ROLL_NO = INTS[0].iloc[7:,1]

def appendtoframe(obj,x):
    write = (pd.DataFrame(x)).iloc[3:,0]
    obj.append(write)

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
GRADES = []

def computeGrades(x):
    ret =  []
    for p in x:
            if(p>=GRADE_RANGES[0]):
                ret.append(3)
            elif(p>=GRADE_RANGES[1]):
                ret.append(2)
            else:
                ret.append(1)
    return ret
def main():
    global MARKS
    global MARKS_PER_QUESTION
    readInternals()
    # print(COS['CO1'])
    # writeToFile(r"./Course Outcomes.xlsx",COS)
    iter = 1

    for flag in COS:
        MARKS.append((pd.DataFrame(COS[flag]).to_numpy()[2:COS_count[flag] + 3,4:]).transpose() )
        MARKS_PER_QUESTION.append((pd.DataFrame(COS[flag]).to_numpy()[2:COS_count[flag] + 3,0]).transpose() )
    # SUM = None
    for l,k in zip(MARKS_PER_QUESTION,MARKS) :

        l = np.array(l,dtype=float)
        k = np.array(k,dtype=float)
        SUM = np.nansum(k,axis=1)

        ATTEMPTED = ~np.isnan(k)
        ATTEMPTED_SUM =np.sum(np.multiply(ATTEMPTED,l),axis=1)
        PERCENTAGE = np.divide(SUM,ATTEMPTED_SUM) * 100
        GRADES = computeGrades(PERCENTAGE)

        SUM = np.append([np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan],SUM)
        ATTEMPTED_SUM = np.append([np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan],ATTEMPTED_SUM)
        PERCENTAGE = np.append([np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan],PERCENTAGE)
        GRADES = np.append([np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan],GRADES)
        appendtoframe(COS['CO' + str(iter)], SUM)
        appendtoframe(COS['CO' + str(iter)], ATTEMPTED_SUM)
        appendtoframe(COS['CO' + str(iter)], PERCENTAGE)
        appendtoframe(COS['CO' + str(iter)], GRADES)
        # print(COS['CO' + str(iter)])
        iter+=1
        # print(PERCENTAGE)

    for key in COS:
        COS[key][COS_count[key]+2][6] = 'Total obtained'
        COS[key][COS_count[key]+3][6] = 'Total Attempted'
        COS[key][COS_count[key]+4][6] = 'Percentage'
        COS[key][COS_count[key]+5][6] = 'Grades on scale of 3'

    writeToFile(r"./Course Outcomes.xlsx",COS)



if( __name__ == "__main__"):
    main()
