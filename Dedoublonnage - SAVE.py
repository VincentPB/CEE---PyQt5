#Importe un fichier .xlsx puis détecte l'opération à appliquer
#dessus pour ensuite le traiter.

#=============================== IMPORTS =================================#

import sys
import os
import os.path
import xlrd
import xlsxwriter
import datetime
#import pyexcel as pe
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from xlwt import Workbook
#from xlutils.copy import copy 
from xlrd import open_workbook
from functools import partial
from tkinter import filedialog

#=============================== GLOBAL =================================#

Temps = datetime.datetime.now()
address = ''

#=============================== FUNCTIONS =================================#

def factorisation(L, Lind):
    l = len(L)
    doubles=[[0]]*l
    indices=[]
    for i in range(l):
        dejaPresent=False
        for j in range(l):
            if (L[i][0] == doubles[j][0]):
                dejaPresent = True
        if(dejaPresent==False):
            doubles[i]=L[i]
            indices.append(Lind[i])

    A=[]
    for i in doubles:
        if i!=[0]:
            A.append(i)
            
    LF=[A,indices]
    return LF

def importer(lbl1, button): #Importe un fichier et en affiche le titre
    global address
    titre = openFileNameDialog()
    titreF=os.path.basename(titre)
    address = titre
    lbl1.setText(titreF)

def traitement(lbl1): #Lance le traitement du fichier
    global address
    if(address!=''):
        switchOperation(address)
        showDialog()

#=========================== OPERATION TREATMENT ============================#

def switchOperation(filename):
    
    wbNew = xlrd.open_workbook(filename)
    OperationName = wbNew.sheet_names()[0]
    print(OperationName)

    if (("TRA-EQ-115" in OperationName) or ("TRA-EQ-15" in OperationName)):
        processEQ115(filename)
        print('TRAITEMENT TERMINE')
    elif (("TRA-EQ-119" in OperationName) or ("TRA-EQ-19" in OperationName)):
        processEQ119(filename)
        print('TRAITEMENT TERMINE')
    elif (("TRA-EQ-103" in OperationName) or ("TRA-EQ-03" in OperationName)):
        processEQ103(filename)
        print('TRAITEMENT TERMINE')
    elif (("TRA-EQ-101" in OperationName) or ("TRA-EQ-01" in OperationName)):
        processEQ101(filename)
        print('TRAITEMENT TERMINE')
    elif (("TRA-EQ-111" in OperationName) or ("TRA-EQ-11" in OperationName)):
        processEQ111(filename)
        print('TRAITEMENT TERMINE')
    elif (("TRA-SE-101" in OperationName) or ("TRA-SE-01" in OperationName)):
        processSE101(filename)
        print('TRAITEMENT TERMINE')
    elif (("TRA-SE-113" in OperationName) or ("TRA-SE-13" in OperationName)):
        processSE113(filename)
        print('TRAITEMENT TERMINE')
    else:
        return("OPERATION INVALIDE")

#=============================== PROCESSING =================================#

        #---------------------- EQ 115 ----------------------#

def processEQ115(filename):

    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax = datetime.datetime(Temps.year-10, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        if(len(str(sheetNew.cell_value(i, 0)))!=0):
            Immati = sheetNew.cell_value(i, 0)
            ListeImmat.append(Immati)

            Datei = xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, 1), wbNew.datemode)

            ListeDate.append(Datei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')

    for j in range(len(ListeImmat)):

        if(ListeDate[j]<TempsMax):
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            Current = ListeImmat[j]
            for k in range(len(ListeImmat)):

                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                    ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS

                    ListeIndiceDoublon.append(k+6)
                
    fueillasse = PostTra.add_worksheet('TRA EQ 115')
    doublons = PostTra.add_worksheet('Doublons')
    V = factorisation(ListeDoublon, ListeIndiceDoublon)
    ListeFinaleDoublons = V[0]
    ListeIndiceDoublonF= V[1]

    ligneD=0
    for i in range(len(ListeFinaleDoublons)):
        for ind in range(9):
            if (ind ==1 or ind==2 or ind==7):
                q = str(xlrd.xldate.xldate_as_datetime(int(str(ListeFinaleDoublons[i][ind])[:-2]), wbNew.datemode))[:10]
                Y = q[8:10]+'/'+q[5:7]+'/'+q[:4]
                doublons.write(ligneD,ind,Y)
            else:
                doublons.write(ligneD,ind,str(ListeFinaleDoublons[i][ind]))
        ligneD+=1
     
     
    ligne=6
    for h1 in range(6):
        for h2 in range(9):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    for i in range(6,len(ListeImmat)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(9):
                if (ind ==1 or ind==2 or ind==7): #ind==8 normalement
                    p = str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10]
                    X = p[8:10]+'/'+p[5:7]+'/'+p[:4]
                    fueillasse.write(ligne, ind, X)
                elif (ind==4):
                    fueillasse.write(ligne, ind, str(sheetNew.row_values(i)[ind])[:-2])
                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1
    #print(ListeImmat)
    PostTra.close()

        #---------------------- EQ 119 ----------------------#

def processEQ119(filename):
    
    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax103 = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(5,NbLines):
        if(sheetNew.cell_value(i, 2)!='' and sheetNew.cell_value(i, 1)!=''):
            Immati = sheetNew.cell_value(i, 2)
            #print(str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 7))[:-2]), wbNew.datemode))[:10])
            Datei = xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 1))[:-2]), wbNew.datemode)
            ListeImmat.append(Immati)
            ListeDate.append(Datei)
        

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    for j in range(len(ListeDate)):
        if(ListeDate[j]<TempsMax103):
            ListeDoublon.append(sheetNew.row_values(j+5)) #DOUBLONS
            ListeIndiceDoublon.append(j+5)
        else:
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+5)) #DOUBLONS
                     ListeIndiceDoublon.append(k+5)
               
    fueillasse = PostTra.add_worksheet('TRA EQ 119')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(11):
                if (ind ==1):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=5
    for h1 in range(5):
        for h2 in range(11):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(5,len(ListeImmat)+5):
        if (i not in ListeIndiceDoublonF):
            for ind in range(11):
                if (ind ==1):
                    a=str(sheetNew.row_values(i)[ind])[:-2]
                    if(len(a)==5):
                        fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
                    else:
                        fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
                elif(ind==5):
                    fueillasse.write(ligne, ind, str(sheetNew.row_values(i)[ind]))
                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()

        #---------------------- EQ 103 ----------------------#

def processEQ103(filename):
    wbNew = xlrd.open_workbook(filename)
    sheetNew = wbNew.sheet_by_index(0)
    if(sheetNew.cell_value(1, 2)!=''):
        TITLE = sheetNew.cell_value(1, 2)
    else:
        TITLE = sheetNew.cell_value(1, 3)
    if('SERIE' in TITLE):
        processEQ103Serie(filename)
        
    elif('INTERNE' in TITLE):
        processEQ103INT(filename)
        
    elif('EXTERNE' in TITLE):
        processEQ103EXT(filename)
        
    else:
        print('FICHIER INVALIDE')

    
        #---------------------- EQ 103 Serie ----------------------#

def processEQ103Serie(filename):

    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax103 = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        if(sheetNew.cell_value(i, 1)!='' and sheetNew.cell_value(i, 7)!=''):
            Immati = sheetNew.cell_value(i, 1)
            #print(str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 7))[:-2]), wbNew.datemode))[:10])
            Datei = xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 7))[:-2]), wbNew.datemode)
            ListeImmat.append(Immati)
            ListeDate.append(Datei)
        

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    for j in range(len(ListeDate)):
        if(ListeDate[j]<TempsMax103):
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                     ListeIndiceDoublon.append(k+6)
               
    fueillasse = PostTra.add_worksheet('TRA EQ 103 Serie')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(9):
                if (ind ==7 or ind==8):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=6
    for h1 in range(6):
        for h2 in range(9):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(6,len(ListeImmat)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(9):
                if (ind ==7 or ind==8):
                    a=str(sheetNew.row_values(i)[ind])[:-2]
                    if(len(a)==5):
                        fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
                    else:
                        fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])

                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()

        #---------------------- EQ 103 EXT ----------------------#

def processEQ103EXT(filename):

    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax103 = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        if(sheetNew.cell_value(i, 2)!='' and sheetNew.cell_value(i, 8)!=''):
            Immati = sheetNew.cell_value(i, 2)
            #print(str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 7))[:-2]), wbNew.datemode))[:10])
            Datei = xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 8))[:-2]), wbNew.datemode)
            ListeImmat.append(Immati)
            ListeDate.append(Datei)
        

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    for j in range(len(ListeDate)):
        if(ListeDate[j]<TempsMax103):
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                     ListeIndiceDoublon.append(k+6)
               
    fueillasse = PostTra.add_worksheet('TRA EQ 103 EXT')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(10):
                if (ind ==9 or ind==8):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=6
    for h1 in range(6):
        for h2 in range(10):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(6,len(ListeImmat)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(10):
                if (ind ==9 or ind==8):
                    a=str(sheetNew.row_values(i)[ind])[:-2]
                    if(len(a)==5):
                        fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
                    else:
                        fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])

                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()

        #---------------------- EQ 103 INT ----------------------#

def processEQ103INT(filename):

    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax103 = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        if(sheetNew.cell_value(i, 1)!='' and sheetNew.cell_value(i, 7)!=''):
            Immati = sheetNew.cell_value(i, 1)
            #print(str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 7))[:-2]), wbNew.datemode))[:10])
            Datei = xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, 7))[:-2]), wbNew.datemode)
            ListeImmat.append(Immati)
            ListeDate.append(Datei)
        

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    for j in range(len(ListeDate)):
        if(ListeDate[j]<TempsMax103):
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                     ListeIndiceDoublon.append(k+6)
               
    fueillasse = PostTra.add_worksheet('TRA EQ 103 INT')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(9):
                if (ind ==7 or ind==8):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=6
    for h1 in range(6):
        for h2 in range(9):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(6,len(ListeImmat)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(9):
                if (ind ==7 or ind==8):
                    a=str(sheetNew.row_values(i)[ind])[:-2]
                    if(len(a)==5):
                        fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
                    else:
                        fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])

                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()
    
        #---------------------- EQ 101 ----------------------#

def processEQ101(filename):
    
    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax101 = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        
        Immati = sheetNew.cell_value(i, 1)
        ListeImmat.append(Immati)
        Datei = xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, 7), wbNew.datemode)
        ListeDate.append(Datei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    
    for j in range(len(ListeImmat)):
        if(ListeDate[j]<TempsMax101):
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            Current = ListeImmat[j]
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                     ListeIndiceDoublon.append(k+6)
                 
    fueillasse = PostTra.add_worksheet('TRA EQ 101')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(9):
                if (ind ==4 or ind==5 or ind==7 or ind==8):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=6
    for h1 in range(6):
        for h2 in range(9):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(6,len(ListeImmat)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(9):
                if (ind ==4 or ind==5 or ind==7 or ind==8):
                    a=str(sheetNew.row_values(i)[ind])[:-2]
                    if(len(a)==5):
                        fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
                    else:
                        fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])

                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()

        #---------------------- SE 111 ----------------------#

def processEQ111(filename):
    
    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax101 = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        
        Immati = sheetNew.cell_value(i, 0)
        ListeImmat.append(Immati)
        Datei = xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, 6), wbNew.datemode)
        ListeDate.append(Datei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    
    for j in range(len(ListeImmat)):
        if(ListeDate[j]<TempsMax101):
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            Current = ListeImmat[j]
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                     ListeIndiceDoublon.append(k+6)
                 
    fueillasse = PostTra.add_worksheet('TRA EQ 111')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(8):
                if (ind ==6 or ind==7):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=6
    for h1 in range(6):
        for h2 in range(8):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(6,len(ListeImmat)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(8):
                if (ind ==6 or ind==7):
                    a=str(sheetNew.row_values(i)[ind])[:-2]
                    if(len(a)==5):
                        fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
                    else:
                        fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])

                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()

        #---------------------- SE 101 ----------------------#

def processSE101(filename):

    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax101 = datetime.datetime(Temps.year-3, Temps.month, Temps.day, 0, 0, 0)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeNom=[]
    ListePrenom=[]
    ListeSIRET=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):

        Nomi = sheetNew.cell_value(i, 1)
        Prenomi = sheetNew.cell_value(i, 0)
        SIRETi = sheetNew.cell_value(i, 3)
        if(Nomi != '' and SIRETi != '' and Prenomi != ''):
            ListeSIRET.append(SIRETi)
            ListeNom.append(Nomi)
            ListePrenom.append(Prenomi)
        if(sheetNew.cell_value(i, 7)!='' and len(str(sheetNew.cell_value(i, 7)))==10):

            Datei = datetime.datetime(int(sheetNew.cell_value(i, 7)[6:]), int(sheetNew.cell_value(i, 7)[3:5]), int(sheetNew.cell_value(i, 7)[0:2]),0,0,0)
            ListeDate.append(Datei)
        elif (sheetNew.cell_value(i, 7)!='' and len(str(sheetNew.cell_value(i, 7))[:-2])==5):
            Datei = xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, 7), wbNew.datemode)
            ListeDate.append(Datei)
        else:
            Datei = datetime.datetime(1, 1, 1,0,0,0)
            ListeDate.append(Datei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    
    for j in range(len(ListeNom)):
        if(ListeDate[j]<TempsMax101):
            #print(ListeDate[j])
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            Current = ListeNom[j]
            for k in range(len(ListeNom)):
                if (k!=j and ListeNom[k]==ListeNom[j] and ListePrenom[k]==ListePrenom[j] and ListeSIRET[k]==ListeSIRET[j]):
                    ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                    ListeIndiceDoublon.append(k+6)
                                  
    fueillasse = PostTra.add_worksheet('TRA SE 101')
    doublons = PostTra.add_worksheet('Doublons')

    V = factorisation(ListeDoublon, ListeIndiceDoublon)
    ListeFinaleDoublons = V[0]
    ListeIndiceDoublonF= V[1]

    ligneD=0
    for i in range(len(ListeFinaleDoublons)):
        for ind in range(9):
            if (ind == 6 or ind == 7):
                #q = str(xlrd.xldate.xldate_as_datetime(int(str(ListeFinaleDoublons[i][ind])[:-2]), wbNew.datemode))[:10]
                q = ListeFinaleDoublons[i][ind]
                #Y = q[8:10]+'/'+q[5:7]+'/'+q[:4]
                doublons.write(ligneD,ind,q)
                        
            else:
                doublons.write(ligneD,ind,str(ListeFinaleDoublons[i][ind]))
        ligneD+=1

   
    ligne=6
    for h1 in range(6):
        for h2 in range(8):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])
    

    for i in range(6,len(ListeNom)):
        if (i not in ListeIndiceDoublonF):
            for ind in range(8):
                a=sheetNew.row_values(i)[ind]
                #if (type(a)==type(0.1) and len(str(a))==5 and (ind == 7 or ind == 8) and len(str(sheetNew.row_values(i)[ind])!=0)):
                if(ind==6 or ind==7):
                        #print(int(str(sheetNew.row_values(i)[ind])[:-2]))
                    #Xi=str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10]
                    Xi=sheetNew.cell_value(i, ind)
                    #Xi=datetime.datetime(int(sheetNew.cell_value(i, ind)[6:]), int(sheetNew.cell_value(i, ind)[3:5]), int(sheetNew.cell_value(i, ind)[0:2]),0,0,0)
                    #Xi=xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, ind), wbNew.datemode)
    
                    #Di=Xi[0:4]+'/'+Xi[5:7]+'/'+Xi[8:10]
                    #print(Xi)
                    fueillasse.write(ligne, ind, Xi)
                        
                #elif(type(a)==type('yolo') and (ind == 7 or ind == 8) and len(sheetNew.row_values(i)[ind])!=0):
                    #fueillasse.write(ligne, ind, str(datetime.date(int(sheetNew.cell_value(i, 7)[6:]), int(sheetNew.cell_value(i, 7)[3:5]), int(sheetNew.cell_value(i, 7)[0:2])))[:10])

                elif (ind == 3):
                    fueillasse.write(ligne, ind, str(sheetNew.row_values(i)[ind])[:-2])
                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])

            ligne+=1

    PostTra.close()

        #---------------------- SE 113 ----------------------#

def processSE113(filename):
    
    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax101 = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        
        Immati = sheetNew.cell_value(i, 0)
        ListeImmat.append(Immati)
        Datei = xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, 2), wbNew.datemode)
        ListeDate.append(Datei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    
    for j in range(len(ListeImmat)):
        if(ListeDate[j]<TempsMax101):
            ListeDoublon.append(sheetNew.row_values(j+6)) #DOUBLONS
            ListeIndiceDoublon.append(j+6)
        else:
            Current = ListeImmat[j]
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                     ListeIndiceDoublon.append(k+6)
                 
    fueillasse = PostTra.add_worksheet('TRA EQ 113')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(7):
                if (ind ==2):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=6
    for h1 in range(6):
        for h2 in range(7):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(6,len(ListeImmat)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(7):
                if (ind ==2):
                    a=str(sheetNew.row_values(i)[ind])[:-2]
                    if(len(a)==5):
                        fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
                    else:
                        fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])

                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()

#=========================== DISPLAY FUNCTION ============================#

def showDialog(): #PopUp de fin de traitement
    msgBox = QMessageBox()
    msgBox.setGeometry(500,350, 200, 200)
    msgBox.setText("<p align='center'>Le dédoublonnage a été effectué avec succès </p>")
    msgBox.setWindowTitle("Traitement terminé")
    msgBox.setFont(QFont("Calibri", 11, QFont.Bold))
    msgBox.setStyleSheet(
    "QPushButton {"
    " font: bold 14px;"
    " min-width: 10em;"
    " padding: 3px;"
    " margin-right:4.5em;"
    "}"
    "* {"
    " margin-right:1.8em;"
    "min-width: 22em;"
    "}"
    );
    msgBox.exec()

def aProposDe(): #PopUp 'A propos'
    msgBox = QMessageBox()
    msgBox.setGeometry(510,350, 200, 200)
    msgBox.setText("<p align='center'>Cette application est une propriété</p> \n <p align='center'>Stela Produits Pétroliers</p>")
    msgBox.setWindowTitle("A propos")
    msgBox.setFont(QFont("Calibri", 11, QFont.Bold))
    msgBox.setStyleSheet(
    "QPushButton {"
    " font: bold 14px;"
    " min-width: 10em;"
    " padding: 3px;"
    " margin-right:3.5em;"
    "}"
    "* {"
    " margin-right:1.8em;"
    "min-width: 20em;"
    "}"
    );
    msgBox.exec()
  
def openFileNameDialog(): #Retourne le nom du fichier sélectionné
        fileName = QFileDialog.getOpenFileName()
        return fileName[0]

class MyMainWindow(QMainWindow): #Fenêtre

    def __init__(self, parent=None):

        super(MyMainWindow, self).__init__(parent)
        self.form_widget = Example(self) 
        self.setCentralWidget(self.form_widget)
        self.setGeometry(450, 250, 500, 250)
        self.setWindowTitle('Dédoublonnage')
        self.setWindowIcon(QIcon('stela.ico'))

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('Options')
        
        exitButton = QAction(QIcon('exit24.png'), 'Quitter', self)
        exitButton.setShortcut('Ctrl+Q')
        exitButton.setStatusTip("Quitter l'application")
        exitButton.triggered.connect(self.close)
        aPropos = QAction(QIcon('exit24.png'), 'A propos', self)
        aPropos.triggered.connect(aProposDe)

        fileMenu.addAction(aPropos)
        fileMenu.addAction(exitButton)

class Example(QWidget): #Widget
    
    def __init__(self, parent):
        super(Example, self).__init__(parent)
        self.initUI()
        
    def initUI(self):

        buttonI = QPushButton('IMPORTER', self)
        buttonI.setToolTip('Importer un fichier à traiter')
        buttonI.clicked.connect(lambda : importer(lbl1, buttonI))
        buttonI.move(75, 150)
        buttonI.setFont(QFont("Calibri", 12, QFont.Bold))
        buttonI.resize(150, 50)

        buttonT = QPushButton('TRAITEMENT', self)
        buttonT.setToolTip('Lancer le traitement du fichier')
        buttonT.clicked.connect(lambda : traitement(lbl1))
        buttonT.move(275, 150)
        buttonT.setFont(QFont("Calibri", 12, QFont.Bold))
        buttonT.resize(150, 50)

        lbl1 = QLabel('Sélectionnez un fichier à importer', self)
        lbl1.setFont(QFont("Calibri", 14, QFont.Bold))
        lbl1.setAlignment(Qt.AlignCenter)
        lbl1.setWordWrap(True)
        lbl1.setGeometry(75,10, 350, 100)

        self.show()

#================================ DISPLAY =================================#
        
app = QApplication([])

        #----------------- STYLE DARK ----------------------#

app.setStyle('Fusion')  
dark_palette = QPalette()
dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
dark_palette.setColor(QPalette.WindowText, Qt.white)
dark_palette.setColor(QPalette.Base, QColor(25, 25, 25))
dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
dark_palette.setColor(QPalette.ToolTipText, Qt.white)
dark_palette.setColor(QPalette.Text, Qt.white)
dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
dark_palette.setColor(QPalette.ButtonText, Qt.white)
dark_palette.setColor(QPalette.BrightText, Qt.red)
dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
dark_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
dark_palette.setColor(QPalette.HighlightedText, Qt.black)
app.setPalette(dark_palette)
app.setStyleSheet("QToolTip { color: #ffffff; background-color: #2a82da; border: 1px solid white; }")

        #----------------- AFFICHAGE ----------------------#

foo = MyMainWindow()
foo.show()
sys.exit(app.exec_())
