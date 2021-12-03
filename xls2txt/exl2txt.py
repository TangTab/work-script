#coding=utf-8

import xlrd
import numpy as np
import sys
import os

def printUsage():
    print ('''usage: exl2txt.py <input excel path> <output path> <number of ncols>
    <input excel path>\t\t necessary, Path of the excel file to be converted
    <output path>\t\t\t optional, Def value is "./output/"
    <number of ncols>\t\t optional, Def value is 5
    ''')

def open_excel(file= 'file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

def excel_to_txt(excel_file, save_path, ncols):
    x1 = open_excel(excel_file)
    for i in x1.sheet_names():
        table = x1.sheet_by_name(i)
        nrows = table.nrows
        data_matrix = np.zeros((nrows, ncols),dtype='object')    # def ncols is 5
        for i in range(ncols):
            cols = ['-' if x == '' else x for x in table.col_values(i)]
            data_matrix[:, i] = cols
        np.savetxt(str(save_path)+str(table.name)+".txt",data_matrix,'%s',delimiter='\t')

if __name__=="__main__":
    n_par = len(sys.argv)
    if n_par >= 2:
        save_path = "./output/"
        ncols = 5
        excel_file = str(sys.argv[1])
        if n_par >= 3:
            save_path = str(sys.argv[2])
        if n_par == 4:
            ncols = int(sys.argv[3])
        if n_par > 4:
            printUsage()
            exit(0)
        if save_path[-1] != '/':
            save_path += '/'
        try:
            os.mkdir(save_path)
        except Exception as e:
            pass
        excel_to_txt(excel_file, save_path, ncols)
    else:
        printUsage()
        exit(0)
