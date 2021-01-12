import pandas as pd
import numpy as np
from itertools import combinations
import xlrd
import pyxlsb
import datetime as dt
from math import comb

from multiprocessing import Pool

# enter the desired number of processes here
NUM_PROCS = 24

global_index_list = []


# Clean month
def policyWithMonthDot(policy):
    condition = [
        policy.str.contains('.JAN'),
        policy.str.contains('.FEB'),
        policy.str.contains('.MAR'),
        policy.str.contains('.APR'),
        policy.str.contains('.MAY'),
        policy.str.contains('.JUN'),
        policy.str.contains('.JUL'),
        policy.str.contains('.AUG'),
        policy.str.contains('.SEP'),
        policy.str.contains('.OCT'),
        policy.str.contains('.NOV'),
        policy.str.contains('.DEC'),
        policy.str.contains('.EN') & (~policy.str.startswith('EN'))]

    output = [
        policy.str.find('.JAN'),
        policy.str.find('.FEB'),
        policy.str.find('.MAR'),
        policy.str.find('.APR'),
        policy.str.find('.MAY'),
        policy.str.find('.JUN'),
        policy.str.find('.JUL'),
        policy.str.find('.AUG'),
        policy.str.find('.SEP'),
        policy.str.find('.OCT'),
        policy.str.find('.NOV'),
        policy.str.find('.DEC'),
        policy.str.find('.EN')]
    return [a[:b] for a, b in zip(policy, np.select(condition, output, policy.str.len()))]


def policyWithMonth(policy):
    condition = [
        policy.str.contains('JAN'),
        policy.str.contains('FEB'),
        policy.str.contains('MAR'),
        policy.str.contains('APR'),
        policy.str.contains('MAY'),
        policy.str.contains('JUN'),
        policy.str.contains('JUL'),
        policy.str.contains('AUG'),
        policy.str.contains('SEP'),
        policy.str.contains('OCT'),
        policy.str.contains('NOV'),
        policy.str.contains('DEC'),
        policy.str.contains('EN') & (~policy.str.startswith('EN'))]

    output = [
        policy.str.find('JAN'),
        policy.str.find('FEB'),
        policy.str.find('MAR'),
        policy.str.find('APR'),
        policy.str.find('MAY'),
        policy.str.find('JUN'),
        policy.str.find('JUL'),
        policy.str.find('AUG'),
        policy.str.find('SEP'),
        policy.str.find('OCT'),
        policy.str.find('NOV'),
        policy.str.find('DEC'),
        policy.str.find('EN')]
    return [a[:b] for a, b in zip(policy, np.select(condition, output, policy.str.len()))]


# Function to create an iterable, number that creates smaller combination first
def clean_policy(x):
    if x[-1] in ('.', '-'):
        return x[:-1]
    else:
        return x


def right_find(policyCol, cpcAmountCol, accAmountCol):
    condition = [cpcAmountCol != 0 and accAmountCol == 0 and policyCol.rfind('.') > policyCol.rfind('-'),
                 cpcAmountCol != 0 and accAmountCol == 0 and policyCol.rfind('.') < policyCol.rfind('-')]

    result = [policyCol[0:policyCol.rfind('.')],
              policyCol[0:policyCol.rfind('-')]]
    return np.select(condition, result, policyCol)


def create_iterable(length):
    x = [i for i in range(length, 1, -1)]
    y = [i for i in range(2, length + 1)]
    z = list(zip(x, y))
    return [x for y in z for x in y][0:len(z)]


def check_negative(iterable):
    lst = [x for x in iterable if x < 0]
    if len(lst) == 1 and sum(iterable) < -1000:  # If sum of a combination < limit ==> no net-off amount
        return False
    else:
        return True


def netOffCheckNonMarine(x, df, col, totalLimit, combinationLimit):
    # col : columns to select
    # threshold : amount limit
    # global combinationList
    global global_index_list
    separateGroup = df.loc[x.index, col]
    separateGroup = separateGroup.sort_values(col)
    dict = separateGroup.to_dict()
    index_list = []
    print('separateGroup:Checking...')
    print('Number of rows: ', len(separateGroup))
    print(separateGroup.head(1))
    print('\n')

    for i in create_iterable(len(dict[col[0]])):
        print('len of dict:', len(dict[col[0]]))
        print('Combination:', i)
        if len(dict[col[0]]) > 1 and check_negative(list(dict[col[1]].values())) == True:
            if (i > len(dict[col[0]])) or comb(len(dict[col[0]].keys()), i) > combinationLimit:
                continue
            index_com = combinations(dict[col[0]].keys(), i)
            print('combination of:', len(dict[col[0]].keys()), i)
            print(dict[col[0]].keys())
            amount_com = combinations(dict[col[1]].values(), i)
            for j, k in zip(amount_com, index_com):
                # if first item of each combination > 0=> no net-off ( df is sorted ascending)
                if j[0] > 0:
                    break
                # if sum of an array=0 and each array does not contain an item of another array
                if np.absolute(np.sum(np.array(j))) <= totalLimit and len(
                        [x for x in k if x in [y for z in index_list for y in z]]) == 0:
                    index_list.append(k)
                    global_index_list.append(k)
                    print('Sum=0 found:', k)
                    # combinationList.append(k) # add the combination to list
                    for l in k:
                        del dict[col[0]][l]
                        del dict[col[1]][l]

    index_list = [x for y in index_list for x in y]
    index_list.sort()
    index_list2 = ["Net-off" if i in index_list else "" for i in x.index]
    print('Done!')
    return index_list2


def process_single_df(df):
    colNonMarine = ['PolicyForChecking', 'Chênh lệch VND']
    df['Check'] = df.groupby(['PolicyForChecking'], sort=False)[['Chênh lệch VND']].transform(netOffCheckNonMarine,
                                                                                              df=df, col=colNonMarine,
                                                                                              totalLimit=10,
                                                                                              combinationLimit=10000000)
    return df


if __name__ == '__main__':
    df = pd.read_excel('D:/GWP.xlsx')
    print('reading done!')

    import time

    start_time = time.time()

    df['Số đơn'] = df['Số đơn'].str.strip()
    df['Số đơn đến mã phòng'] = df['Số đơn đến mã phòng'].str.strip()
    df['Số đơn'] = df['Số đơn'].apply(clean_policy)
    df['Số đơn đến mã phòng'] = df['Số đơn đến mã phòng'].apply(clean_policy)
    df['PolicyForChecking'] = df.apply(
        lambda x: right_find(x['Số đơn'], x['Phí bảo hiểm VND CPC'], x['Phí bảo hiểm VND ACC']), axis=1).astype(str)
    df['PolicyForChecking'] = policyWithMonthDot(df['PolicyForChecking'])
    df['PolicyForChecking'] = policyWithMonth(df['PolicyForChecking'])

    dfok = df.loc[(df['Chênh lệch NT nhỏ nhất'].abs() < 1) | (df['Chênh lệch VND nhỏ nhất'] == 0)]
    dfcheck = df.loc[(df['Chênh lệch NT nhỏ nhất'].abs() >= 1) & (df['Chênh lệch VND nhỏ nhất'] != 0)]
    dfList = [dfg for _, dfg in dfcheck.groupby(['PolicyForChecking'])]

    pool = Pool(processes=NUM_PROCS)

    allDfs = pool.map(process_single_df, dfList)
    print('Process done!')
    f = open("D:/time.txt", "a")
    f.write(str((time.time() - start_time)))
    f.close()
    allDfs.append(dfok)
    dfFinal = pd.concat(allDfs, axis=0)
    dfFinal.to_excel('D:/df.xlsx')
    print('done!')
