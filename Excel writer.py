import pandas as pd
import numpy as np
from fileWrite import *
import ntpath
import easygui
import os
import warnings
warnings.filterwarnings("ignore")


# array['CO1'][0] = 8
COS = {'CO Summary': [],'CO1': [], 'CO2': [], 'CO3': [], 'CO4': [],
       'CO5': [], 'CO6': [], }
CO_SUMMARY = [[None for f in range(100)] for x in range(100)]

INT_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
ENDSEM_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
GRADE_RANGES = [60, 50, 40]
INTS = []
ENDSEM = None

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

def readFile(file):
    global COS,CO_SUMMARY,INT_COS_count,ENDSEM_COS_count,INTS,ENDSEM,NUM_STUDENTS,INT_MARKS,ENDSEM_MARKS
    global INT_MARKS_PER_QUESTION,ENDSEM_MARKS_PER_QUESTION,QUESTION_ATTEMPTED,GRADES,EMPTY,INT_NUM_GRADES,ENDSEM_NUM_GRADES
    global INT_AVERAGE_GRADE
    global ENDSEM_AVERAGE_GRADE
    COS = {'CO Summary': [],'CO1': [], 'CO2': [], 'CO3': [], 'CO4': [],
        'CO5': [], 'CO6': [], }
    CO_SUMMARY = [[None for f in range(100)] for x in range(100)]

    INT_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
    ENDSEM_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
    INTS = []
    ENDSEM = None

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
    try:
        print('Reading file: ',file)
        # Creating file pointer in the name of excel
        excel = pd.ExcelFile(file)
        # Array INTS holds the objects for all internals sheets of excelfile as its elements
        INTS = [pd.read_excel(excel, 'INT1'), pd.read_excel(excel, 'INT2')]
        ENDSEM = pd.read_excel(excel, 'End Sem')
        NAMES = INTS[0].iloc[3:, 2]  # TODO
        ROLL_NO = INTS[0].iloc[3:, 1]
    except:
        print('Error reading file : ',ntpath.basename(file))


# User defined function


def initSummary():
    startRow = 3
    startColumn = 3
    iter = 1
    CO_SUMMARY[startColumn + 3][startRow - 2 +
                                        (iter - 1) * 7] =  'No of\nStudents'
    for key in COS:
        if(key != 'CO Summary'):
            CO_SUMMARY[startColumn + 4][startRow +
                                        (iter - 1) * 7] = 'CO' + str(iter)
            CO_SUMMARY[startColumn - 1][startRow + 2 +
                                        (iter - 1) * 7] = 'Number of students'
            CO_SUMMARY[startColumn - 1][startRow +
                                        4 + (iter - 1) * 7] = 'Percentage'

            CO_SUMMARY[startColumn][startRow + 1 +
                                    (iter - 1) * 7] = 'Total INT\nGrade 3'
            CO_SUMMARY[startColumn + 1][startRow + 1 +
                                        (iter - 1) * 7] = 'Total INT\nGrade 2'
            CO_SUMMARY[startColumn + 2][startRow + 1 +
                                        (iter - 1) * 7] = 'Total INT\nGrade 1'
            CO_SUMMARY[startColumn + 3][startRow +
                                        1 + (iter - 1) * 7] = 'INT Avg\nGrade'

            CO_SUMMARY[startColumn + 5][startRow + 1 +
                                        (iter - 1) * 7] = 'Total ES\nGrade 3'
            CO_SUMMARY[startColumn + 6][startRow + 1 +
                                        (iter - 1) * 7] = 'Total ES\nGrade 2'
            CO_SUMMARY[startColumn + 7][startRow + 1 +
                                        (iter - 1) * 7] = 'Total ES\nGrade 1'
            CO_SUMMARY[startColumn + 8][startRow +
                                        1 + (iter - 1) * 7] = 'ES Avg\nGrade'

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
    


# print(COS)



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


def addToSummary(entry, iter, t, g,type='num'):
    startRow = 3
    startColumn = 3
    TEST = {'int':0,'end':5}
    GRD = {'3':0,'2':1,'1':2,'avg':3}
    type_offset = 1
    if (type == 'percentage'):
        type_offset = 3
    CO_SUMMARY[startColumn + TEST[t] + GRD[g]][startRow + 1 +
                                        type_offset + (iter - 1) * 7] = entry


def main(OP_FILE):
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
        if(INT_COS_count['CO' + str(iter)] != 0):
            appendtoframe(COS['CO' + str(iter)], SUM)
            appendtoframe(COS['CO' + str(iter)], ATTEMPTED_SUM)
            appendtoframe(COS['CO' + str(iter)], PERCENTAGE)
            appendtoframe(COS['CO' + str(iter)], GRADES)
        else:
            for i in range(4):
                appendtoframe(COS['CO' + str(iter)], EMPTY)
        # END SEM
        # if(ENDSEM_NUM_GRADES != )
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
        if(ENDSEM_COS_count['CO' + str(iter)] != 0):
            appendtoframe(COS['CO' + str(iter)], SUM)
            appendtoframe(COS['CO' + str(iter)], ATTEMPTED_SUM)
            appendtoframe(COS['CO' + str(iter)], PERCENTAGE)
            appendtoframe(COS['CO' + str(iter)], GRADES)
        else:
            for i in range(4):
                appendtoframe(COS['CO' + str(iter)], EMPTY)

        appendtoframe(COS['CO' + str(iter)], EMPTY)
        appendtoframe(COS['CO' + str(iter)], EMPTY)  # Spacing
        appendtoframe(COS['CO' + str(iter)], EMPTY)

        if(INT_COS_count['CO' + str(iter)] != 0):
            appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['3'])
            addToSummary(INT_NUM_GRADES['3'][8],iter, 'int', '3')
            addToSummary(INT_NUM_GRADES['3'][10],iter, 'int', '3','percentage')

            appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['2'])
            addToSummary(INT_NUM_GRADES['2'][8],iter, 'int', '2')
            addToSummary(INT_NUM_GRADES['2'][10],iter, 'int', '2','percentage')

            appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['1'])
            addToSummary(INT_NUM_GRADES['1'][8],iter, 'int', '1')
            addToSummary(INT_NUM_GRADES['1'][10],iter, 'int', '1','percentage')

            appendtoframe(COS['CO' + str(iter)], INT_AVERAGE_GRADE)
            addToSummary(INT_AVERAGE_GRADE[8],iter, 'int', 'avg')
            addToSummary(INT_AVERAGE_GRADE[8],iter, 'int', 'avg','percentage')
        else:
            addToSummary('NA',iter, 'int', '3')
            addToSummary('NA',iter, 'int', '3','percentage')
            addToSummary('NA',iter, 'int', '2')
            addToSummary('NA',iter, 'int', '2','percentage')
            addToSummary('NA',iter, 'int', '1')
            addToSummary('NA',iter, 'int', '1','percentage')
            addToSummary('NA',iter, 'int', 'avg')
            addToSummary('NA',iter, 'int', 'avg','percentage')
            
            for i in range(4):
                appendtoframe(COS['CO' + str(iter)], EMPTY)


        if(ENDSEM_COS_count['CO' + str(iter)] != 0):
            appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['3'])
            addToSummary(ENDSEM_NUM_GRADES['3'][8],iter, 'end', '3')
            addToSummary(ENDSEM_NUM_GRADES['3'][10],iter, 'end', '3','percentage')

            appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['2'])
            addToSummary(ENDSEM_NUM_GRADES['2'][8],iter, 'end', '2')
            addToSummary(ENDSEM_NUM_GRADES['2'][10],iter, 'end', '2','percentage')

            appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['1'])
            addToSummary(ENDSEM_NUM_GRADES['1'][8],iter, 'end', '1')
            addToSummary(ENDSEM_NUM_GRADES['1'][10],iter, 'end', '1','percentage')

            appendtoframe(COS['CO' + str(iter)], ENDSEM_AVERAGE_GRADE)
            addToSummary(ENDSEM_AVERAGE_GRADE[8],iter, 'end', 'avg')
            addToSummary(ENDSEM_AVERAGE_GRADE[10],iter, 'end', 'avg','percentage')
        else:
            addToSummary('NA',iter, 'end', '3')
            addToSummary('NA',iter, 'end', '3','percentage')
            addToSummary('NA',iter, 'end', '2')
            addToSummary('NA',iter, 'end', '2','percentage')
            addToSummary('NA',iter, 'end', '1')
            addToSummary('NA',iter, 'end', '1','percentage')
            addToSummary('NA',iter, 'end', 'avg')
            addToSummary('NA',iter, 'end', 'avg','percentage')
            
            for i in range(4):
                appendtoframe(COS['CO' + str(iter)], EMPTY)


        iter += 1
        # print(COS['CO' + str(iter)])
        # print(PERCENTAGE)
    print('Number of questions: ')
    for key in COS:
        if(key != 'CO Summary'):
            print('\t    ',key)
            print('Internals: %d\tEnd Semester: %d '%(INT_COS_count[key],ENDSEM_COS_count[key]))
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
    CO_SUMMARY[7][1] = NUM_STUDENTS
    COS['CO Summary'] = CO_SUMMARY
    writeToFile(OP_FILE, COS)



if(__name__ == "__main__"):
    files = easygui.fileopenbox("Excel writer", "Choose excel files to process", filetypes= ['*',"*.xlsx"], multiple=True) 
    string = 'Do you want to CONTINUE with same grade ranges?\n\n\nCurrent Ranges are:\n\n Grade 3 Minimum Percentage : ' + str(GRADE_RANGES[0]) + '\n Grade 2 Minimum Percentage : ' + str(GRADE_RANGES[1]) #+' \n Grade 1 Minimum Percentage : ' + str(GRADE_RANGES[2])
    while(not easygui.ynbox(string,title='Continue with same grade ranges?')):
        GRADE_RANGES[0] = easygui.integerbox(title='Enter Min Percentage for Grade 3')
        GRADE_RANGES[1] = easygui.integerbox(title='Enter Min Percentage for Grade 2',upperbound=GRADE_RANGES[0])
        # GRADE_RANGES[2] = easygui.integerbox(title='Enter Min Percentage for Grade 1',upperbound=GRADE_RANGES[1])

        string = 'Do you want to CONTINUE with same grade ranges?\n\n\nCurrent Ranges are:\n\n Grade 3 Minimum Marks : ' + str(GRADE_RANGES[0]) + '\n Grade 2 Minimum Marks : ' + str(GRADE_RANGES[1]) #+' \n Grade 1 Minimum Marks : ' + str(GRADE_RANGES[2])

    for file in files:
        readFile(file)
        OP_FILE = os.path.join(ntpath.split(file)[0],'Summary ' + ntpath.basename(file))

        try:
            main(OP_FILE)
        except:
            print('Please check contents of file : ',ntpath.basename(file))
            os.system('pause ')

    os.system('pause ')
