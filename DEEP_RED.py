import sys
import os
import os.path
import xlrd
import xlsxwriter
import datetime
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from xlwt import Workbook
from xlrd import open_workbook
from functools import partial
from tkinter import filedialog

def isDate(x):
    if(type(x)==float and len(str(x))==7):
        return True
    return False

filename='DONE/TR_TRA-EQ-111.xlsx'
wbNew = xlrd.open_workbook(filename)
sheetNew = wbNew.sheet_by_index(0)
print(isDate(sheetNew.cell_value(8, 6)))





def info(filename):
    wbNew = xlrd.open_workbook(filename)
    sheetNew = wbNew.sheet_by_index(0)
    nbLines = sheetNew.nrows

    IndicesDate=[]
    Bool = True
    for k in range(6):
        ligne = sheetNew.row_values(k)
        if Bool:
            j=0
            while(j<len(ligne) and ligne[j]!=''):
                if ("DATE" in ligne[j] or 'Date' in ligne[j]):
                    Bool=False
                    IndicesDate.append(j)
                    lineStart = k+1
                j+=1
            lineLen=j

    B=False
    lastLine=nbLines
    for i in range(lineStart, nbLines):
        if(len(sheetNew.cell_value(i, 0))==0 and B==False):
            lastLine = i
            B=True

    NBLINES = lastLine
    NBCOL = lineLen
    DEBUT = lineStart
    INDDATE = IndicesDate

    for h in range(DEBUT):
        for t in range(NBCOL):
            for l in range(len(sheetNew.cell_value(h, t))-3):
                if('TRA' in sheetNew.cell_value(h, t)[l:l+3]):
                    OPERATION = sheetNew.cell_value(h, t)[l+4:]
            
    return(OPERATION, NBLINES, NBCOL, DEBUT, INDDATE)
