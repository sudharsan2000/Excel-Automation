import pandas as pd
import numpy as np
import ntpath
import easygui
import os
import warnings
import time
import math

from openpyxl import load_workbook


warnings.filterwarnings("ignore")


# array['CO1'][0] = 8
CO_percentage = {'CO1': [], 'CO2': [], 'CO3': [], 'CO4': [],
       'CO5': [], 'CO6': [], }
POS = {}
PO = None
POS_count= {}
POS_write= []
COS = {'CO Summary': [],'CO1': [], 'CO2': [], 'CO3': [], 'CO4': [],
       'CO5': [], 'CO6': [], }
CO_SUMMARY = [[None for f in range(100)] for x in range(100)]

INT_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
ENDSEM_COS_count = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
showFlagCOS = {'CO1': 0, 'CO2': 0, 'CO3': 0, 'CO4': 0, 'CO5': 0, 'CO6': 0}
GRADE_RANGES = [60, 50, 40]
INTS = []
ENDSEM = None
x_overall = {}
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
overallPercentage = {}
INT_AVERAGE_GRADE = []
ENDSEM_AVERAGE_GRADE = []
sheet_to_df_map = {}

def writeToFile(path,original,obj,trans=1):
    global showFlagCOS
    print('Writing to file : ',path)
    writer = pd.ExcelWriter(path,mode = "w") 
    for key in original:
        k = original[key]
        #write = pd.DataFrame(k)
        k.to_excel(writer,key,header=False,index=False) 
    for key in obj:

        if(key == 'CO Summary' or key == 'Percentage of COs POs and PSOs'):
            if (trans):
                k = map(list,zip(*obj[key]))
            else:
                k = obj[key]
            write = pd.DataFrame(k)
            write.to_excel(writer,key,header=False,index=False)

        elif (showFlagCOS[key] != 0 ):
            if (trans):
                k = map(list,zip(*obj[key]))
            else:
                k = obj[key]
            write = pd.DataFrame(k)
            write.to_excel(writer,key,header=False,index=False)
    writer.save()
def readFile(file):
    global COS,CO_SUMMARY,INT_COS_count,ENDSEM_COS_count,INTS,ENDSEM,NUM_STUDENTS,INT_MARKS,ENDSEM_MARKS
    global INT_MARKS_PER_QUESTION,ENDSEM_MARKS_PER_QUESTION,QUESTION_ATTEMPTED,GRADES,EMPTY,INT_NUM_GRADES,ENDSEM_NUM_GRADES
    global INT_AVERAGE_GRADE,PO
    global ENDSEM_AVERAGE_GRADE,sheet_to_df_map
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
    sheet_to_df_map = {}
    PO = None
    
    try:
        print('Reading file: ',file)
        # Creating file pointer in the name of excel
        excel = pd.ExcelFile(file)
        for sheet_name in excel.sheet_names:
            sheet_to_df_map[sheet_name] = excel.parse(sheet_name)
        # Array INTS holds the objects for all internals sheets of excelfile as its elements
        for key, value in sheet_to_df_map.items():
            if (key == 'PO-Attainment'):
                PO = value
            else:
                INTS.append(value)            
                if (key == 'End Sem'):
                    ENDSEM = value
        NAMES = INTS[0].iloc[3:, 2]  
        ROLL_NO = INTS[0].iloc[3:, 1]
        readPOS()
        excel.close()
    except:
        print('Error reading file : ',ntpath.basename(file))


# User defined function
def readPOS():
    global POS, POS_count
    for i in range(1, 1000):
        try:
            s = PO.iat[1, 1+i]
            if (s.startswith('P')):
                POS[str(s)] = PO.iloc[2:, 1 + i]
                POS_count[str(s)] = 1 
            else:
                pass
        except:
            break
def calculatePOS():
    global POS,POS_count,POS_write,CO_percentage,EMPTY,x_overall
    a = np.append([np.nan, np.nan],INTS[0].iloc[6:, 1])  # Roll Number
    POS_write.append(a)
    a = np.append([np.nan, np.nan],INTS[0].iloc[6:, 2]) # Name 
    POS_write.append(a)

    EMPTY = np.ones(a.shape) * np.nan
    POS_write.append(EMPTY)
    POS_write.append(EMPTY)

    ## Append CO percentages here
    for i in range(1,7):
            try:
                # a = np.append([np.nan, np.nan, np.nan],CO_percentage['CO' + str(i)])
                # a = concat( ['', '', 'CO' +  str(i)], np.array_str(CO_percentage['CO' + str(i)]) )
                a =  ['', '', 'CO' +  str(i) + ' %'] + ["%.2f" % x if (math.isnan(x) == 0) else ' ' for x in CO_percentage['CO' + str(i)]] 
                POS_write.append(a)
            except:
                pass

    POS_write.append(EMPTY)
    PO_number = 1
    for p in POS:
        x = 0
        count = 0
        x_overall[p] = 0
        for i in range(1,7):
            try:
                x += CO_percentage['CO' + str(i)] * POS[p][i + 1]
                x_overall[p] += overallPercentage['CO' + str(i)] * POS[p][i + 1]

                count += POS[p][i+1]
            except:
                pass
        x = x / count
        x_overall[p] = truncate(x_overall[p] / count)
        if (PO_number <= 12):
            x = ['', '', 'PO' +  str(PO_number) + ' %' ] + ["%.2f" % itera if (math.isnan(itera) == 0) else ' ' for itera in x]
        else:
            x = ['', '', 'PSO' +  str(PO_number - 12) + ' %' ] + ["%.2f" % itera if (math.isnan(itera) == 0) else ' ' for itera in x]

        POS_write.append(x)
        PO_number = PO_number + 1

def initSummary():
    startRow = 3
    startColumn = 3
    iter = 1
    CO_SUMMARY[startColumn + 3][startRow - 3 +
                                        (iter - 1) * 7] =  'Total Students'

    for key in COS:
        if(key != 'CO Summary'):
            CO_SUMMARY[startColumn + 4][startRow - 1 +
                                        (iter - 1) * 7] = 'CO' + str(iter)
            CO_SUMMARY[startColumn - 2][startRow + 2 +
                                        (iter - 1) * 7] = 'Number of students'
            CO_SUMMARY[startColumn - 2][startRow + 
                                        4 + (iter - 1) * 7] = 'Percentage'
            
            CO_SUMMARY[startColumn + 1][startRow  +
                                    (iter - 1) * 7] = 'Cumulative'
            CO_SUMMARY[startColumn][startRow + 1 +
                                    (iter - 1) * 7] = 'Grade 3'
            CO_SUMMARY[startColumn + 1][startRow + 1 +
                                        (iter - 1) * 7] = 'Grade 2'
            CO_SUMMARY[startColumn + 2][startRow + 1 +
                                        (iter - 1) * 7] = 'Grade 1'
            CO_SUMMARY[startColumn + 3][startRow +
                                        1 + (iter - 1) * 7] = 'Avg Grade'

            CO_SUMMARY[startColumn + 6][startRow  +
                                        (iter - 1) * 7] = 'End Semester'
            CO_SUMMARY[startColumn + 5][startRow + 1 +
                                        (iter - 1) * 7] = 'Grade 3'
            CO_SUMMARY[startColumn + 6][startRow + 1 +
                                        (iter - 1) * 7] = 'Grade 2'
            CO_SUMMARY[startColumn + 7][startRow + 1 +
                                        (iter - 1) * 7] = 'Grade 1'
            CO_SUMMARY[startColumn + 8][startRow  +
                                        1 + (iter - 1) * 7] = 'Avg Grade'
            
            
            iter += 1
    i = 0 
    CO_SUMMARY[startColumn + 4 + i][startRow - 2 +
                                        (iter ) * 7] = 'Overall POs'
    for p in POS:    
        CO_SUMMARY[startColumn -2 + i][startRow - 1 +
                                        (iter ) * 7] = p
        i +=1
def finishSummary():
    global x_overall,CO_SUMMARY
    startColumn = 3
    startRow = 3
    i = 0
    iter = 7
    for p in POS:    
        CO_SUMMARY[startColumn -2 + i][startRow +
                                        (iter ) * 7] = x_overall[p]
        i +=1
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
        print('Finding COs from sheet :', iter)
        iter += 1
        for i in range(1, 1000):
            try:
                COS[str(INT.iat[5, 2+i])].append(INT.iloc[3:, 2+i])
                # print(INT.iat[5,2+i])
                INT_COS_count[INT.iat[5, 2+i]] += 1
            except:
                break
    #print('Finding COs from ENDSEM')
    for i in range(1, 1000):
        try:
            ENDSEM_COS_count[ENDSEM.iat[5, 2+i]] += 1
            # COS[str(ENDSEM.iat[5, 2+i])].append(ENDSEM.iloc[3:, 2+i])
            EMPTY = np.ones(ENDSEM.iloc[3:, 2+i].shape) * np.nan
            COS[str(ENDSEM.iat[5, 2+i])].append(EMPTY)
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
    global NUM_STUDENTS, overallPercentage,x_overall
    readInternals()
    initSummary()
    # print(COS['CO1'])
    # writeToFile(r"./Course Outcomes.xlsx",COS)
    iter = 1

    for flag in COS:
        if(flag != 'CO Summary'):
            INT_MARKS.append((pd.DataFrame(COS[flag]).to_numpy()[
                2:INT_COS_count[flag] + 2, 4:]).transpose())
            INT_MARKS_PER_QUESTION.append((pd.DataFrame(COS[flag]).to_numpy()[
                2:INT_COS_count[flag] + 2, 0]).transpose())
                ########################################################################################################################
            ENDSEM_MARKS.append((pd.DataFrame(COS[flag]).to_numpy()[
                                INT_COS_count[flag] - ENDSEM_COS_count[flag] + 2:INT_COS_count[flag] + 2, 4:]).transpose())
            ENDSEM_MARKS_PER_QUESTION.append((pd.DataFrame(COS[flag]).to_numpy(
            )[INT_COS_count[flag] - ENDSEM_COS_count[flag] + 2:INT_COS_count[flag] + 2, 0]).transpose())

    for l, k, m, n in zip(INT_MARKS_PER_QUESTION, INT_MARKS, ENDSEM_MARKS_PER_QUESTION, ENDSEM_MARKS):
        # INT

        l = np.array(l, dtype=float)
        k = np.array(k, dtype=float)
        SUM = np.nansum(k, axis=1)

        ATTEMPTED = ~np.isnan(k)
        ATTEMPTED_SUM = np.sum(np.multiply(ATTEMPTED, l), axis=1)
        PERCENTAGE = (np.divide(SUM, ATTEMPTED_SUM) * 100).round(decimals=2)
        CO_percentage['CO' + str(iter)] = PERCENTAGE
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
            for i in range(4):            ##################################################################################################################################
                    appendtoframe(COS['CO' + str(iter)], EMPTY)

        appendtoframe(COS['CO' + str(iter)], EMPTY)
        appendtoframe(COS['CO' + str(iter)], EMPTY)  # Spacing
        appendtoframe(COS['CO' + str(iter)], EMPTY)

        if(INT_COS_count['CO' + str(iter)] != 0):
            # appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['3'])
            addToSummary(INT_NUM_GRADES['3'][8],iter, 'int', '3')
            addToSummary(INT_NUM_GRADES['3'][10],iter, 'int', '3','percentage')
            overallPercentage['CO' + str(iter)] = INT_NUM_GRADES['3'][10]

            # appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['2'])
            addToSummary(INT_NUM_GRADES['2'][8],iter, 'int', '2')
            addToSummary(INT_NUM_GRADES['2'][10],iter, 'int', '2','percentage')

            # appendtoframe(COS['CO' + str(iter)], INT_NUM_GRADES['1'])
            addToSummary(INT_NUM_GRADES['1'][8],iter, 'int', '1')
            addToSummary(INT_NUM_GRADES['1'][10],iter, 'int', '1','percentage')

            # appendtoframe(COS['CO' + str(iter)], INT_AVERAGE_GRADE)
            addToSummary(INT_AVERAGE_GRADE[8],iter, 'int', 'avg')
        else:
            addToSummary('NA',iter, 'int', '3')
            addToSummary('NA',iter, 'int', '3','percentage')
            addToSummary('NA',iter, 'int', '2')
            addToSummary('NA',iter, 'int', '2','percentage')
            addToSummary('NA',iter, 'int', '1')
            addToSummary('NA',iter, 'int', '1','percentage')
            addToSummary('NA',iter, 'int', 'avg')
            
            for i in range(4):
                appendtoframe(COS['CO' + str(iter)], EMPTY)


        if(ENDSEM_COS_count['CO' + str(iter)] != 0):    
            # appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['3'])
            addToSummary(ENDSEM_NUM_GRADES['3'][8],iter, 'end', '3')
            addToSummary(ENDSEM_NUM_GRADES['3'][10],iter, 'end', '3','percentage')

            # appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['2'])
            addToSummary(ENDSEM_NUM_GRADES['2'][8],iter, 'end', '2')
            addToSummary(ENDSEM_NUM_GRADES['2'][10],iter, 'end', '2','percentage')

            # appendtoframe(COS['CO' + str(iter)], ENDSEM_NUM_GRADES['1'])
            addToSummary(ENDSEM_NUM_GRADES['1'][8],iter, 'end', '1')
            addToSummary(ENDSEM_NUM_GRADES['1'][10],iter, 'end', '1','percentage')

            # appendtoframe(COS['CO' + str(iter)], ENDSEM_AVERAGE_GRADE)
            addToSummary(ENDSEM_AVERAGE_GRADE[8],iter, 'end', 'avg')
        else:
            addToSummary('NA',iter, 'end', '3')
            addToSummary('NA',iter, 'end', '3','percentage')
            addToSummary('NA',iter, 'end', '2')
            addToSummary('NA',iter, 'end', '2','percentage')
            addToSummary('NA',iter, 'end', '1')
            addToSummary('NA',iter, 'end', '1','percentage')
            addToSummary('NA',iter, 'end', 'avg')
            
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
            showFlagCOS[key] = INT_COS_count[key] + ENDSEM_COS_count[key]
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     2][5] = 'CUMULATIVE'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     2][6] = 'Obtained'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     3][6] = 'Attempted'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 4][6] = 'Percentage'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     5][6] = 'Grades'


            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 6][5] = 'END SEMESTER ONLY'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 6][6] = 'Obtained'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 7][6] = 'Attempted'
            COS[key][INT_COS_count[key] +
                     ENDSEM_COS_count[key] + 8][6] = 'Percentage'
            COS[key][INT_COS_count[key] + ENDSEM_COS_count[key] +
                     9][6] = 'Grades'

    calculatePOS()       
    CO_SUMMARY[8][0] = NUM_STUDENTS
    finishSummary()
    COS['CO Summary'] = CO_SUMMARY
    COS['Percentage of COs POs and PSOs'] = POS_write
    writeToFile(OP_FILE,sheet_to_df_map, COS)



if(__name__ == "__main__"):
    files = easygui.fileopenbox("Excel writer", "Choose excel files to process", filetypes= ["*.xlsx",'*'], multiple=True) 
    string = 'Do you want to CONTINUE with same grade ranges?\n\n\nCurrent Ranges are:\n\n Grade 3 Minimum Percentage : ' + str(GRADE_RANGES[0]) + '\n Grade 2 Minimum Percentage : ' + str(GRADE_RANGES[1]) #+' \n Grade 1 Minimum Percentage : ' + str(GRADE_RANGES[2])
    while(not easygui.ynbox(string,title='Continue with same grade ranges?')):
        temp =  easygui.integerbox(title='Enter Min Percentage for Grade 3',default=GRADE_RANGES[0])
        GRADE_RANGES[0] = temp if temp != None else GRADE_RANGES[0]
        temp = easygui.integerbox(title='Enter Min Percentage for Grade 2',upperbound=GRADE_RANGES[0],default=GRADE_RANGES[1])
        GRADE_RANGES[1] = temp if temp != None else GRADE_RANGES[1]
        # GRADE_RANGES[2] = easygui.integerbox(title='Enter Min Percentage for Grade 1',upperbound=GRADE_RANGES[1])

        string = 'Do you want to CONTINUE with same grade ranges?\n\n\nCurrent Ranges are:\n\n Grade 3 Minimum Marks : ' + str(GRADE_RANGES[0]) + '\n Grade 2 Minimum Marks : ' + str(GRADE_RANGES[1]) #+' \n Grade 1 Minimum Marks : ' + str(GRADE_RANGES[2])

    for file in files:
        tick = time.time()
        readFile(file)
        OP_FILE = os.path.join(ntpath.split(file)[0],'Output ' + ntpath.basename(file))
        # OP_FILE = file

        try:
            
            main(OP_FILE)
            print(time.time() - tick)
        except:
            print('Please check contents of file : ',ntpath.basename(file))
            print('Or close the Output excel files.')
            os.system('pause ')

    os.system('pause ')
