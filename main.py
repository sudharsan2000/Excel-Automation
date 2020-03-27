import pandas as pd
import numpy as np
from fileWrite import *

# array['CO1'][0] = 8
COS = {'CO1': [], 'CO2': [], 'CO3': [], 'CO4': [],
       'CO5': [], 'CO6': [], 'CO Summary': []}
CO_SUMMARY = [ [ None for f in range(100)] for x in range( 100) ]

INT_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
ENDSEM_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
GRADE_RANGES = [60, 50, 40]
try:
    # Creating file pointer in the name of excel
    excel = pd.ExcelFile('CO-PO MOS U15AET502 -odd 2018-19-Micro.xlsx')
    # Array INTS holds the objects for all internals sheets of excelfile as its elements
    INTS = [pd.read_excel(excel, 'INT1'), pd.read_excel(excel, 'INT2')]
    ENDSEM = pd.read_excel(excel, 'End Sem')
except:
    print('Error reading file')
    exit(1)

NAMES = INTS[0].iloc[3:, 2]  # TODO
ROLL_NO = INTS[0].iloc[3:, 1]

# User defined function

def initSummary():
    startRow = 3
    startColumn = 3 
    iter = 1
    for key in COS:
        if(key != 'CO Summary'):
            CO_SUMMARY[startColumn + 4 ][startRow + (iter -1 ) * 7] = 'CO' + str(iter)
            CO_SUMMARY[startColumn - 1 ][startRow + 2 + (iter -1 ) * 7] = 'Number of students'
            CO_SUMMARY[startColumn - 1 ][startRow + 4 + (iter -1 ) * 7] = 'Percentage'
            iter += 1 

def appendtoframe(obj, x, series=0):

    write = (pd.DataFrame(x)).iloc[3:, 0]
    if(series == 0):
        obj.append(write)
    else:
        obj.append(pd.Series(x))


def readInternals():
    iter = 1
    COS['CO1'].append(INTS[0].iloc[3:, 1])
    COS['CO1'].append(INTS[0].iloc[3:, 2])

    COS['CO2'].append(INTS[0].iloc[3:, 1])
    COS['CO2'].append(INTS[0].iloc[3:, 2])

    COS['CO3'].append(INTS[0].iloc[3:, 1])
    COS['CO3'].append(INTS[0].iloc[3:, 2])

    COS['CO4'].append(INTS[0].iloc[3:, 1])
    COS['CO4'].append(INTS[0].iloc[3:, 2])

    COS['CO5'].append(INTS[0].iloc[3:, 1])
    COS['CO5'].append(INTS[0].iloc[3:, 2])

    COS['CO6'].append(INTS[0].iloc[3:, 1])
    COS['CO6'].append(INTS[0].iloc[3:, 2])

    for INT in INTS:
        print('Finding COs from INT', iter)
        iter += 1
        for i in range(1, 1000):
            try:
                COS[str(INT.iat[5, 2+i])].append(INT.iloc[3:, 2+i])
                # print(INT.iat[5,2+i])
                INT_COS_count[INT.iat[5, 2+i]] += 1
            except:
                break
    print('Finding COs from ENDSEM')
    for i in range(1, 1000):
        try:
            COS[str(ENDSEM.iat[5, 2+i])].append(ENDSEM.iloc[3:, 2+i])
            ENDSEM_COS_count[ENDSEM.iat[5, 2+i]] += 1
        except:
            break

    print('Done!')
    print('Number of CO1 questions found : ', INT_COS_count['CO1'])
    print('Number of CO2 questions found : ', INT_COS_count['CO2'])
    print('Number of CO3 questions found : ', INT_COS_count['CO3'])
    print('Number of CO3 questions found : ', INT_COS_count['CO4'])
    print('Number of CO3 questions found : ', INT_COS_count['CO5'])
    print('Number of CO3 questions found : ', INT_COS_count['CO6'])


# print(COS)
NUM_STUDENTS = 0
INT_MARKS = []
ENDSEM_MARKS = []
INT_MARKS_PER_QUESTION = []
ENDSEM_MARKS_PER_QUESTION = []
QUESTION_ATTEMPTED = []
GRADES = []
EMPTY = []
INT_NUM_GRADES = {'1': 0, '2': 0, '3': 0}
ENDSEM_NUM_GRADES = {'1': 0, '2': 0, '3': 0}

INT_AVERAGE_GRADE = []
ENDSEM_AVERAGE_GRADE = []


def truncate(f, n=2):
    '''Truncates/pads a float f to n decimal places without rounding'''
    s = '{}'.format(f)
    if 'e' in s or 'E' in s:
        return '{0:.{1}f}'.format(f, n)
    i, p, d = s.partition('.')
    return '.'.join([i, (d+'0'*n)[:n]])


def computeGrades(x):
    ret = []
    for p in x:
        if(p >= GRADE_RANGES[0]):
            ret.append(3)
        elif(p >= GRADE_RANGES[1]):
            ret.append(2)
        else:
            ret.append(1)
    return ret

def addToSummary(a,b,c):
    pass

def main():
    global INT_MARKS
    global ENDSEM_MARKS
    global INT_MARKS_PER_QUESTION
    global ENDSEM_MARKS_PER_QUESTION
    global INT_NUM_GRADES
    global ENDSEM_NUM_GRADES
    global NUM_STUDENTS
    readInternals()
    initSummary()
    # print(COS['CO1'])
    # writeToFile(r"./Course Outcomes.xlsx",COS)
    iter = 1

    for flag in COS:
        if(flag != 'CO Summary'):
            INT_MARKS.append((pd.DataFrame(COS[flag]).to_numpy()[
                2:INT_COS_count[flag] + 3, 4:]).transpose())
            INT_MARKS_PER_QUESTION.append((pd.DataFrame(COS[flag]).to_numpy()[
                2:INT_COS_count[flag] + 3, 0]).transpose())
            ENDSEM_MARKS.append((pd.DataFrame(COS[flag]).to_numpy()[
                                INT_COS_count[flag] + 2:INT_COS_count[flag] + ENDSEM_COS_count[flag] + 4, 4:]).transpose())
            ENDSEM_MARKS_PER_QUESTION.append((pd.DataFrame(COS[flag]).to_numpy(
            )[INT_COS_count[flag] + 2:INT_COS_count[flag] + ENDSEM_COS_count[flag] + 4, 0]).transpose())

    for l, k, m, n in zip(INT_MARKS_PER_QUESTION, INT_MARKS, ENDSEM_MARKS_PER_QUESTION, ENDSEM_MARKS):
        # INT

        l = np.array(l, dtype=float)
        k = np.array(k, dtype=float)
        SUM = np.nansum(k, axis=1)

        ATTEMPTED = ~np.isnan(k)
        ATTEMPTED_SUM = np.sum(np.multiply(ATTEMPTED, l), axis=1)
        PERCENTAGE = (np.divide(SUM, ATTEMPTED_SUM) * 100).round(decimals=2)
        GRADES = computeGrades(PERCENTAGE)

        NUM_STUDENTS = np.size(GRADES)

        SUM = np.append([np.nan, np.nan, np.nan, np.nan,
                         np.nan, np.nan, np.nan], SUM)
        ATTEMPTED_SUM = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], ATTEMPTED_SUM)
        PERCENTAGE = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], PERCENTAGE)
        EMPTY = np.ones(PERCENTAGE.shape) * np.nan

        INT_NUM_GRADES['1'] = EMPTY.copy()
        INT_NUM_GRADES['2'] = EMPTY.copy()
        INT_NUM_GRADES['3'] = EMPTY.copy()

        ENDSEM_NUM_GRADES['1'] = EMPTY.copy()
        ENDSEM_NUM_GRADES['2'] = EMPTY.copy()
        ENDSEM_NUM_GRADES['3'] = EMPTY.copy()

        INT_AVERAGE_GRADE = EMPTY.copy()
        ENDSEM_AVERAGE_GRADE = EMPTY.copy()

        GRADE_LIST = list(GRADES)
        INT_NUM_GRADES['1'][1] = GRADE_LIST.count(1)
        INT_NUM_GRADES['1'][3] = truncate(
            (GRADE_LIST.count(1)/NUM_STUDENTS) * 100)

        INT_NUM_GRADES['2'][1] = GRADE_LIST.count(2)
        INT_NUM_GRADES['2'][3] = truncate(
            (GRADE_LIST.count(2)/NUM_STUDENTS) * 100)

        INT_NUM_GRADES['3'][1] = GRADE_LIST.count(3)
        INT_NUM_GRADES['3'][3] = truncate(
            (GRADE_LIST.count(3)/NUM_STUDENTS) * 100)

        INT_AVERAGE_GRADE[1] = np.mean(GRADE_LIST)

        INT_NUM_GRADES['3'] = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], INT_NUM_GRADES['3'])
        INT_NUM_GRADES['2'] = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], INT_NUM_GRADES['2'])
        INT_NUM_GRADES['1'] = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], INT_NUM_GRADES['1'])
        INT_AVERAGE_GRADE = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], INT_AVERAGE_GRADE)

        GRADES = np.append([np.nan, np.nan, np.nan, np.nan,
                            np.nan, np.nan, np.nan], GRADES)

        appendtoframe(COS['CO' + str(iter)], SUM)
        appendtoframe(COS['CO' + str(iter)], ATTEMPTED_SUM)
        appendtoframe(COS['CO' + str(iter)], PERCENTAGE)
        appendtoframe(COS['CO' + str(iter)], GRADES)
        # END SEM
# 3
        m = np.array(m, dtype=float)
        n = np.array(n, dtype=float)
        SUM = np.nansum(n, axis=1)

        ATTEMPTED = ~np.isnan(n)
        ATTEMPTED_SUM = np.sum(np.multiply(ATTEMPTED, m), axis=1)
        PERCENTAGE = (np.divide(SUM, ATTEMPTED_SUM) * 100).round(decimals=2)
        GRADES = computeGrades(PERCENTAGE)

        SUM = np.append([np.nan, np.nan, np.nan, np.nan,
                         np.nan, np.nan, np.nan], SUM)
        ATTEMPTED_SUM = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], ATTEMPTED_SUM)
        PERCENTAGE = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], PERCENTAGE)

        GRADE_LIST = list(GRADES)
        ENDSEM_NUM_GRADES['1'][1] = GRADE_LIST.count(1)
        ENDSEM_NUM_GRADES['1'][3] = truncate(
            (GRADE_LIST.count(1)/NUM_STUDENTS) * 100)

        ENDSEM_NUM_GRADES['2'][1] = GRADE_LIST.count(2)
        ENDSEM_NUM_GRADES['2'][3] = truncate(
            (GRADE_LIST.count(2)/NUM_STUDENTS) * 100)

        ENDSEM_NUM_GRADES['3'][1] = GRADE_LIST.count(3)
        ENDSEM_NUM_GRADES['3'][3] = truncate(
            (GRADE_LIST.count(3)/NUM_STUDENTS) * 100)

        ENDSEM_AVERAGE_GRADE[1] = np.mean(GRADE_LIST)

        ENDSEM_NUM_GRADES['3'] = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], ENDSEM_NUM_GRADES['3'])
        ENDSEM_NUM_GRADES['2'] = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], ENDSEM_NUM_GRADES['2'])
        ENDSEM_NUM_GRADES['1'] = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], ENDSEM_NUM_GRADES['1'])
        ENDSEM_AVERAGE_GRADE = np.append(
            [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan], ENDSEM_AVERAGE_GRADE)

        GRADES = np.append([np.nan, np.nan, np.nan, np.nan,
                            np.nan, np.nan, np.nan], GRADES)
        appendtoframe(COS['CO' + str(iter)], SUM)
        appendtoframe(COS['CO' + str(iter)], ATTEMPTED_SUM)
        appendtoframe(COS['CO' + str(iter)], PERCENTAGE)
        appendtoframe(COS['CO' + str(iter)], GRADES)

        appendtoframe(COS['CO' + str(iter)], EMPTY)
        appendtoframe(COS['CO' + str(iter)], EMPTY)  # Spacing
        appendtoframe(COS['CO' + str(iter)], EMPTY)

        appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['3'])
        addToSummary(iter,'int','3')
        appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['2'])
        addToSummary(iter,'int','2')
        appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['1'])
        addToSummary(iter,'int','1')
        appendtoframe(COS['CO' + str(iter)], INT_AVERAGE_GRADE)
        addToSummary(iter,'int','avg')

        appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['3'])
        addToSummary(iter,'end','3')
        appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['2'])
        addToSummary(iter,'end','2')
        appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['1'])
        addToSummary(iter,'end','1')
        appendtoframe(COS['CO' + str(iter)], ENDSEM_AVERAGE_GRADE)
        addToSummary(iter,'end','avg')

        iter += 1
        # print(COS['CO' + str(iter)])
        # print(PERCENTAGE)

    for key in COS:
        if(key != 'CO Summary'):
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     2][6] = 'INT Total obtained'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     3][6] = 'INT Total Attempted'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 4][6] = 'INT Percentage'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     5][6] = 'INT Grades on\nscale of 3'

            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 6][6] = 'END obtained'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 7][6] = 'END Attempted'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 8][6] = 'END Percentage'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     9][6] = 'END Grades on\nscale of 3'

            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] + 12][8] = 'Number:'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 12][10] = 'Percentage:'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     15][14] = 'Number of Students:'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 16][14] = NUM_STUDENTS

            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     13][6] = 'Total INT\nGrade 3'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     14][6] = 'Total INT\nGrade 2'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     15][6] = 'Total INT\nGrade 1'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 16][6] = 'INT\n Avg Grade'

            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     17][6] = 'Total END\nGrade 3'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     18][6] = 'Total END\nGrade 2'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     19][6] = 'Total END\nGrade 1'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 20][6] = 'END\n Avg Grade'
    COS['CO Summary'] = CO_SUMMARY
    writeToFile(r"./Course Outcomes.xlsx", COS)


if(__name__ == "__main__"):
    main()
