# -*- coding: utf-8 -*-
# Import `os`
import os
import pandas as pd
from ahpy import *
import numpy as np

def import_excel():
    # Retrieve current working directory (`cwd`)
    cwd = os.getcwd()
    cwd

    # Change directory
    #os.chdir("/Local/hamedi/data/Python")
    os.chdir(cwd)

    # List all files and directories in current directory
    #os.listdir('.')
    #print(os.listdir('.'))

    # Assign spreadsheet filename to `file`
    file = 'test1.xlsx'

    # Load spreadsheet
    xl = pd.ExcelFile(file)

    # Print the sheet names
    sheets = xl.sheet_names

    # Load a sheet into a DataFrame by name: df1
    df1 = xl.parse(sheets[0]).to_dict('records')

    return df1

def prepare_data_and_compare():
    df = import_excel()
    row1 = list(df[0].keys())
    name = row1[0]
    criterion = row1[1:]
    values = [list(ds.values())[1:] for ds in df]
    cmpr = Compare(name, values, criterion, precision=5, random_index='saaty')
    res= cmpr.report()

    columns = criterion + ["", "Weights"]
    rows = criterion + ["C.R."]


    cmplt=[]
    #print(res[1])

    mtrx_list = res[0].tolist()

    def extnd(orig, ext):
        orig.extend(ext)
        return orig


    for i, rs in enumerate(mtrx_list):

        cmplt.append(extnd(rs,["",res[1][i]]))

    cmplt.append([res[2]])
    print(cmplt)

    return cmplt,columns,rows

def export_excel():

    data,columns,rows = prepare_data_and_compare()
    print(columns)
    print(rows)
    pdd = pd.DataFrame(data,rows,columns)
    print(pdd)
    writer = pd.ExcelWriter('test-out.xlsx', engine='xlsxwriter')

    # Write your DataFrame to a file
    pdd.to_excel(writer, 'AHP-Weights')

    # Save the result
    writer.save()
    return

export_excel()