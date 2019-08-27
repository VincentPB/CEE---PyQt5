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

    if (OperationName=="TRA-EQ-15" or OperationName=="TRA-EQ-115"):
        processEQ115(filename)
        print('TRAITEMENT TERMINE')
    elif (OperationName=="TRA-EQ-114" or OperationName=="TRA-EQ-14"):
        #processEQ114(filename)
        print('TRAITEMENT TERMINE')
    elif (OperationName=="TRA-EQ-103" or OperationName=="TRA-EQ-03"):
        #processEQ103(filename)
        print('TRAITEMENT TERMINE')
    elif (OperationName=="TRA-EQ-101" or OperationName=="TRA-EQ-01"):
        #processEQ101(filename)
        print('TRAITEMENT TERMINE')
    elif (OperationName=="TRA-EQ-111" or OperationName=="TRA-EQ-11"):
        #processEQ111(filename)
        print('TRAITEMENT TERMINE')
    elif (OperationName=="TRA-SE-101" or OperationName=="TRA-SE-01"):
        processSE101(filename)
        print('TRAITEMENT TERMINE')
    elif (OperationName=="TRA-SE-113" or OperationName=="TRA-SE-13"):
        #processSE113(filename)
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

        #---------------------- EQ 114 ----------------------#

def processEQ114(filename):
    return("Opération TRA EQ 114")

        #---------------------- EQ 103 ----------------------#

def processEQ103(filename):

    wbNew = xlrd.open_workbook(filename)
    global Temps
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

    for i in range(1,NbLines):
        
        Immati = sheetNew.cell_value(i, 2)
        ListeImmat.append(Immati)
        if(len(str(sheetNew.cell_value(i, 6)))!=0):
            #print(int(str(sheetNew.cell_value(i, 6))[6:]))
            Datei = datetime.datetime(int(str(sheetNew.cell_value(i, 6))[6:]), int(str(sheetNew.cell_value(i, 6))[3:5]), int(str(sheetNew.cell_value(i, 6))[:2]),0,0,0)
        else:
            Datei = ''
        #Datei = xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, 6), wbNew.datemode)
        ListeDate.append(Datei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    
    for j in range(len(ListeImmat)):
        if(ListeDate[j] != '' and ListeDate[j]<TempsMax):
            ListeDoublon.append(sheetNew.row_values(j+1)) #DOUBLONS
            ListeIndiceDoublon.append(j+1)
        else:
            Current = ListeImmat[j]
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+1)) #DOUBLONS
                     ListeIndiceDoublon.append(k+1)
                 
    fueillasse = PostTra.add_worksheet('TRA EQ 103')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(9):
                if(ind==8 and len(str(ListeDoublon[i][ind]))!=0):
                    p = str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10]
                    X = p[8:10]+'/'+p[5:7]+'/'+p[:4]
                    doublons.write(ligneD,ind,X)
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=1
    for h1 in range(1):
        for h2 in range(9):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(1,len(ListeImmat)+1):
        if (i not in ListeIndiceDoublonF):
            for ind in range(9):
                if (ind ==5 or ind==6):
                    fueillasse.write(ligne, ind, str(sheetNew.row_values(i)[ind]))
                    
                elif(ind==4):
                    fueillasse.write(ligne, ind, str(sheetNew.row_values(i)[ind])[:-2])
                    
                elif(ind==8):
                    if(len(str(sheetNew.row_values(i)[ind])[:-2])!=0):
                        p = str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10]
                        X = p[8:10]+'/'+p[5:7]+'/'+p[:4]
                        fueillasse.write(ligne, ind, X)
                else:
                    fueillasse.write(ligne, ind, sheetNew.row_values(i)[ind])
 
            ligne+=1

    PostTra.close()

    
        #---------------------- EQ 101 ----------------------#

def processEQ101(filename):
    
    wbNew = xlrd.open_workbook(filename)
    global Temps
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
                 
    fueillasse = PostTra.add_worksheet('TRA EQ 101')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(7):
                if (ind ==2 or ind==3 or ind==4 or ind==5):
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
                if (ind ==2 or ind==3 or ind==4 or ind==5):
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
    global Temps
    TempsMax101 = datetime.datetime(Temps.year-10, Temps.month, Temps.day)
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

    for i in range(2,NbLines):
        
        Immati = sheetNew.cell_value(i, 5)
        ListeImmat.append(Immati)
        Datei = xlrd.xldate.xldate_as_datetime(sheetNew.cell_value(i, 16), wbNew.datemode)
        ListeDate.append(Datei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
 
    for j in range(len(ListeImmat)):
        if(ListeDate[j]<TempsMax101):
            ListeDoublon.append(sheetNew.row_values(j+2)) #DOUBLONS
            ListeIndiceDoublon.append(j+2)
        else:
            Current = ListeImmat[j]
            for k in range(len(ListeImmat)):
                if (k!=j and ListeImmat[k]==ListeImmat[j]):
                     ListeDoublon.append(sheetNew.row_values(k+2)) #DOUBLONS
                     ListeIndiceDoublon.append(k+2)
                 
    fueillasse = PostTra.add_worksheet('TRA EQ 111')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(17):
                doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=2
    for h1 in range(2):
        for h2 in range(17):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(2,len(ListeImmat)+2):
        if (i not in ListeIndiceDoublonF):
            for ind in range(17):
                if (ind ==16):
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
    global Temps
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
                    #print('All true')
                    ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                    ListeIndiceDoublon.append(k+6)
                                  
    fueillasse = PostTra.add_worksheet('TRA SE 101')
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    print('BONSOIR')
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(11):
                if(ind==6 or ind==7):
                    D=str(ListeDoublon[i][ind])
                    X=D[0:4]+'/'+D[5:7]+'/'+D[8:10]
                    doublons.write(ligneD,ind,X)
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]
    print('BONSOIR')
    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
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
                    Xi=str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10]
                    Di=Xi[0:4]+'/'+Xi[5:7]+'/'+Xi[8:10]
                    fueillasse.write(ligne, ind, Di)
                    
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
    sheetNew = wbNew.sheet_by_index(0)
    NbLines = sheetNew.nrows
    ListeCartes=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeMatricule=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)  
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    for i in range(6,NbLines):
        
        Cartei = sheetNew.cell_value(i, 0)
        ListeCartes.append(Cartei)
        
        Matriculei = sheetNew.cell_value(i, 3)
        ListeMatricule.append(Matriculei)

    PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
    
    for j in range(len(ListeCartes)):
        Current = ListeCartes[j]
        for k in range(len(ListeCartes)):
            if (k!=j):
                if(ListeCartes[k]==ListeCartes[j] or ListeMatricule[k]==ListeMatricule[j]):
                    ListeDoublon.append(sheetNew.row_values(k+6)) #DOUBLONS
                    ListeIndiceDoublon.append(k+6)

    fueillasse = PostTra.add_worksheet('TRA SE 113')             
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(9):
                if(ind==1):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,ListeDoublon[i][ind])
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])

    ligne=6

    for h1 in range(6):
        for h2 in range(9):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])
    
    for i in range(6, len(ListeCartes)+6):
        if (i not in ListeIndiceDoublonF):
            for ind in range(9):
                if(ind==1):
                    fueillasse.write(ligne, ind, str(xlrd.xldate.xldate_as_datetime(int(str(sheetNew.row_values(i)[ind])[:-2]), wbNew.datemode))[:10])
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
        self.setGeometry(450, 250, 450, 250)
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
        buttonI.move(50, 150)
        buttonI.setFont(QFont("Calibri", 12, QFont.Bold))
        buttonI.resize(150, 50)

        buttonT = QPushButton('TRAITEMENT', self)
        buttonT.setToolTip('Lancer le traitement du fichier')
        buttonT.clicked.connect(lambda : traitement(lbl1))
        buttonT.move(250, 150)
        buttonT.setFont(QFont("Calibri", 12, QFont.Bold))
        buttonT.resize(150, 50)

        lbl1 = QLabel('Sélectionnez un fichier à importer', self)
        lbl1.setFont(QFont("Calibri", 14, QFont.Bold))
        lbl1.setAlignment(Qt.AlignCenter)
        lbl1.setGeometry(60,10, 350, 100)

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
