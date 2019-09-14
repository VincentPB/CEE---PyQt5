#Importe un fichier .xlsx puis détecte l'opération à appliquer
#pour ensuite supprimer les douablons.

#=============================== IMPORTS =================================#

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

#=============================== GLOBAL =================================#

Temps = datetime.datetime.now()
address = ''

#=============================== FUNCTIONS =================================#

def factorisation(L, Lind): #Supprime une des entrées qui apparaissent en double.
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
    lbl1.setText('Vous avez importé : \n\n' + titreF)

def traitement(lbl1): #Lance le traitement du fichier
    global address
    if(address!=''):
        switchOperation(address)
        showDialog()

def info(filename): #Récupère les informations utiles pour le traitement d'un fichier.

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
        if(len(str(sheetNew.cell_value(i, 0)))==0 and B==False):
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

#=========================== OPERATION TREATMENT ============================#

def switchOperation(filename): #Applique le traitement correspondant au fichier importé.
    INFO = info(filename)
    OperationName = INFO[0]

    if(('115+103' in OperationName) or ('103+115' in OperationName)):
        process(filename, INFO, [0, 1], 'TRA EQ 115-103')
        print('TRAITEMENT TERMINE')        
    elif (("EQ-115" in OperationName) or ("EQ-15" in OperationName)):
        process(filename, INFO, [0, 1], 'TRA EQ 115')
        print('TRAITEMENT TERMINE')
    elif (("EQ-119" in OperationName) or ("EQ-19" in OperationName)):
        process(filename, INFO, [2, 1], 'TRA EQ 119')
        print('TRAITEMENT TERMINE')
    elif (("EQ-103" in OperationName) or ("EQ-03" in OperationName)):
        processEQ103(filename, INFO)
        print('TRAITEMENT TERMINE')
    elif (("EQ-101" in OperationName) or ("EQ-01" in OperationName)):
        process(filename, INFO, [1, 7], 'TRA EQ 101')
        print('TRAITEMENT TERMINE')   
    elif (("EQ-111" in OperationName) or ("EQ-11" in OperationName)):
        process(filename, INFO, [0, 6], 'TRA EQ 111')
        print('TRAITEMENT TERMINE')
    elif (("SE-113" in OperationName) or ("SE-13" in OperationName)):
        process(filename, INFO, [1, 3], 'TRA SE 113')
        print('TRAITEMENT TERMINE')
    elif (("SE-108" in OperationName) or ("SE-08" in OperationName)):
        process(filename, INFO, [1, 5], 'TRA SE 108')
        print('TRAITEMENT TERMINE')
    elif (("SE-105" in OperationName) or ("SE-05" in OperationName)):
        process(filename, INFO, [4, 3], 'TRA SE 105')
        print('TRAITEMENT TERMINE')
    elif (("SE-101" in OperationName) or ("SE-01" in OperationName)):
        process(filename, INFO, [2, 1], 'TRA SE 101')
        print('TRAITEMENT TERMINE')
    else:
        return("OPERATION INVALIDE")

#=============================== PROCESSING =================================#

        #---------------------- TOUS ----------------------#

def process(filename, INFO, IND, TITRE): #Traie le document
 
    NbLines = INFO[1]
    NBCOL = INFO[2]
    DEBUT = INFO[3]
    INDDATE = INFO[4]
    wbNew = xlrd.open_workbook(filename)
    Temps = datetime.datetime.now()
    TempsMax = datetime.datetime(Temps.year-12, Temps.month, Temps.day)
    sheetNew = wbNew.sheet_by_index(0)
    ListeImmat=[]
    ListeDoublon=[]
    ListeIndiceDoublon=[]
    ListeDate=[]
    fileDir=os.path.dirname(os.path.realpath(filename))
    NameOfFile=os.path.basename(filename)
    book1 = Workbook()
    feuil1 = book1.add_sheet('Doublons')

    if(TITRE=='TRA SE 101'):
        ListeNom=[]
        ListePrenom=[]
        ListeSIRET=[]
        ListeDoublon=[]
        ListeIndiceDoublon=[]

        for i in range(DEBUT,NbLines):
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

        PostTra = xlsxwriter.Workbook(filename[:-5]+'_DEDOUBLONNE'+'.xlsx')
        for j in range(len(ListeDate)):
            if(ListeDate[j]<TempsMax):
                ListeDoublon.append(sheetNew.row_values(j+DEBUT)) #DOUBLONS
                ListeIndiceDoublon.append(j+DEBUT)
            else:
                for k in range(len(ListePrenom)):
                    if (k!=j and ListePrenom[k]==ListePrenom[j] and ListeNom[k]==ListeNom[j] and ListeSIRET[k]==ListeSIRET[j]):
                         ListeDoublon.append(sheetNew.row_values(k+DEBUT)) #DOUBLONS
                         ListeIndiceDoublon.append(k+DEBUT)

    else:        
        for i in range(DEBUT,NbLines):
            if(sheetNew.cell_value(i, IND[0])!='' and sheetNew.cell_value(i, IND[1])!=''):
                Immati = sheetNew.cell_value(i, IND[0])
                Datei = xlrd.xldate.xldate_as_datetime(int(str(sheetNew.cell_value(i, IND[1]))[:-2]), wbNew.datemode)
                ListeImmat.append(Immati)
                ListeDate.append(Datei)
        

        PostTra = xlsxwriter.Workbook(filename[:-5]+'_POST'+'.xlsx')
        for j in range(len(ListeDate)):
            if(ListeDate[j]<TempsMax):
                ListeDoublon.append(sheetNew.row_values(j+DEBUT)) #DOUBLONS
                ListeIndiceDoublon.append(j+DEBUT)
            else:
                for k in range(len(ListeImmat)):
                    if (k!=j and ListeImmat[k]==ListeImmat[j]):
                         ListeDoublon.append(sheetNew.row_values(k+DEBUT)) #DOUBLONS
                         ListeIndiceDoublon.append(k+DEBUT)
               
    fueillasse = PostTra.add_worksheet(TITRE)
    doublons = PostTra.add_worksheet('Doublons')

    ligneD=0
    
    for i in range(len(ListeDoublon)):
        if i%2 == 0:
            for ind in range(NBCOL):
                if (ind in INDDATE):
                    doublons.write(ligneD, ind, str(xlrd.xldate.xldate_as_datetime(int(str(ListeDoublon[i][ind])[:-2]), wbNew.datemode))[:10])
                else:
                    doublons.write(ligneD,ind,str(ListeDoublon[i][ind]))
            ligneD+=1
    ListeIndiceDoublonF=[]

    for i in range(len(ListeIndiceDoublon)):
        if i%2 == 0:
            ListeIndiceDoublonF.append(ListeIndiceDoublon[i])
            
    ligne=DEBUT
    for h1 in range(DEBUT):
        for h2 in range(NBCOL):
            fueillasse.write(h1, h2, sheetNew.row_values(h1)[h2])

    
    for i in range(DEBUT,len(ListeDate)+DEBUT):
        if (i not in ListeIndiceDoublonF):
            for ind in range(NBCOL):
                if (ind in INDDATE):
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

def processEQ103(filename, INFO): #Sélectionne le bon traitement (Il y a 3 opérations EQ 103 différentes)
    wbNew = xlrd.open_workbook(filename)
    sheetNew = wbNew.sheet_by_index(0)
    if(sheetNew.cell_value(1, 2)!=''):
        TITLE = sheetNew.cell_value(1, 2)
    else:
        TITLE = sheetNew.cell_value(1, 3)
    if('SERIE' in TITLE):
        process(filename, INFO, [1, 7], 'TRA EQ 103 Serie')
        
    elif('INTERNE' in TITLE):
        process(filename, INFO, [2, 8], 'TRA EQ 103 INT')
        
    elif('EXTERNE' in TITLE):
        process(filename, INFO, [1, 7], 'TRA EQ 103 EXT')
        
    else:
        print('FICHIER INVALIDE')

#=========================== DISPLAY FUNCTION ============================#

def showDialog(): #PopUp de fin de traitement
    msgBox = QMessageBox()
    msgBox.setGeometry(475,330, 200, 200)
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
    msgBox.setGeometry(487,330, 200, 200)
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
        self.setGeometry(400, 250, 500, 250)
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
