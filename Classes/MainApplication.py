import tkinter as tk
import os
import json
import xlsxwriter
import xlrd
from openpyxl.workbook import Workbook as openpyxlWorkbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
import shutil
import pprint

from tkinter import *

class MainApplication:



    #Class Variables <------------------------------------->

    univBackColor = "#012e20"
    univForColor = "white"
    univActiveFgColor = "blue"
    univPath = "yello"
    jsonFilePath = "yello"

    #Class Variables <------------------------------------->!



    def __init__(self, master,title, width, height, icon=None, color=None):
        self.master = master
        self.master.geometry(str(width)+"x"+str(height))
        self.master.title(title)
        self.master.iconbitmap(icon)
        self.master.config(bg = color)

        self.universalBackgroundColor = MainApplication.univBackColor
        self.universalForegroundColor = MainApplication.univForColor
        self.universalActiveForegroundColor = MainApplication.univActiveFgColor

        self.buttons = []
        self.labels = []
        self.entries = []
        self.ddmenus = []
        self.ddmenus_clicked = []

        self.nrButtons = 0
        self.nrEntries = 0
        self.nrLabels = 0
        self.nrDdmenus = 0
        self.nrDdmenus_clickeds = 0


    #Widget Creating Methods <------------------------------------->


    def createLabelAtPosition(self,x=0,y=0,text="",paddingX=10,paddingY=10,backgroundColor=None,foregroundColor=None):
        label = Label(self.master,text=text,padx=paddingX,pady=paddingY,bg = self.universalBackgroundColor, fg = self.universalForegroundColor)
        if backgroundColor != None:
            label.config(bg = backgroundColor)
        if foregroundColor != None:
            label.config(fg = foregroundColor)
        label.grid(row = x,column = y)
        self.labels.append(label)
        self.nrLabels = self.nrLabels + 1



    def createEntryAtPosition(self,x=0,y=0,border="3",insertbackgroundColor = None,backgroundColor=None,foregroundColor=None):
        entry = Entry(self.master,bg = self.universalBackgroundColor, fg = self.universalForegroundColor,insertbackground = self.universalForegroundColor,bd = border,exportselection = 0)
        if backgroundColor != None:
            entry.config(bg = backgroundColor)
        if foregroundColor != None:
            entry.config(fg = foregroundColor)
        if insertbackgroundColor != None:
            entry.config(insertbackgournd = insertbackgroundColor)
        entry.grid(row = x,column = y)
        self.entries.append(entry)
        self.nrEntries = self.nrEntries + 1



    def createButtonAtPosition(self,x=0,y=0,text="ok",paddingX=10,paddingY=5,border="3",backgroundColor=None,foregroundColor=None):
        button = Button(self.master,text = text,bg=self.universalBackgroundColor,fg=self.universalForegroundColor,padx=paddingX,pady=paddingY,bd=border)
        if backgroundColor != None:
            button.config(bg = backgroundColor)
        if foregroundColor != None:
            button.config(fg = foregroundColor)
        button.grid(row = x,column = y)
        self.buttons.append(button)
        self.nrButtons = self.nrButtons + 1



    def createDropdownMenuAtPosition(self,options,x=0,y=0,backgroundColor=None,foregroundColor=None):

        clicked = StringVar()
        clicked.set(options[0])

        drop = OptionMenu(self.master, clicked, *options)

        drop.config(bg=self.universalBackgroundColor,fg=self.universalForegroundColor,activebackground = self.universalBackgroundColor, activeforeground = self.universalForegroundColor)
        drop['menu'].config(bg=self.universalBackgroundColor,fg=self.universalForegroundColor,activebackground = self.universalBackgroundColor, activeforeground = self.universalActiveForegroundColor)

        if backgroundColor != None:
            drop['menu'].config(bg = backgroundColor)
        if foregroundColor != None:
            drop['menu'].config(fg = foregroundColor)

        if backgroundColor != None:
            drop.config(bg = backgroundColor,activebackground=backgroundColor)
        if foregroundColor != None:
            drop.config(fg = foregroundColor,activeforeground=foregroundColor)

        drop.grid(row = x,column = y)
        self.ddmenus.append(drop)
        self.nrDdmenus = self.nrDdmenus + 1

        self.ddmenus_clicked.append(clicked)
        self.nrDdmenus_clickeds = self.nrDdmenus_clickeds + 1



    #Widget Creating Methods <------------------------------------->!



    #Widget Management Methods <------------------------------------->

    def destroySelf(self):
        self.master.destroy()

    def getTextFromEntry(self,index=0): #Returns the text from entry 'index'
        text = self.entries[index].get()
        self.entries[index].delete(0)
        return text


    def addNormalCommandToButton(self,index,command): #Adds a simple command to button
        self.buttons[index].config(command = command)


    def addLambdaCommandToButton(self,index,lambda_command,*args): #Adds a command with specifiable parameters to button
        self.buttons[index].config(command = lambda : lambda_command(*args))


    def showSelection(self,index):
        text = self.ddmenus_clicked[index].get()
        print(text)

    def getSelection(self,index):
        text = self.ddmenus_clicked[index].get()
        return text

    #Widget Management Methods <------------------------------------->!







    #Class Methods <------------------------------------->



    @classmethod
    def changeUnivBackColor(cls,color):
        cls.univBackColor = color


    @classmethod
    def changeUnivForColor(cls,color):
        cls.univForColor = color


    @classmethod
    def changeUnivActiveForColor(cls,color):
        cls.univActiveFgColor = color


    @classmethod
    def printPath(cls):
        print(cls.univPath)

    @classmethod
    def initPath(cls,path):
        cls.univPath = path

    @classmethod
    def initJsonPath(cls,path):
        cls.jsonFilePath = path
        cls.jsonFilePath = cls.jsonFilePath + "\\Resources\\paths.json"

    #Class Methods <------------------------------------->!

#File Management Methods Methods <------------------------------------->



#xlrd (for useless very old fcking piece of garbage xls files) OK BOOMER <------------------------------------->

#xlrd (for useless very old fcking piece of garbage xls files) OK BOOMER<------------------------------------->!

#openpyxl (for the modern-ish xlsx files)<------------------------------------->

#openpyxl (for the modern-ish xlsx files)<------------------------------------->!


#other stuff<------------------------------------->

def correctName(name):
    nr = ""
    namae = ""
    for ch in name:
        if ch >= '0' and ch <= '9':
            nr = nr + str(ch)
        else:
            namae = namae + str(ch)

    nr = nr.zfill(2)
    print(namae+nr)
    return namae+nr

def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

def getFileXFromPath(path,x):
    files = []

    orig = os.getcwd()

    os.chdir(path)

    for r,d,f in os.walk("."):
        for file in f:
            if '.xls' in file:
                files.append(os.path.join(r, file))
            elif '.xlsx' in file:
                files.append(os.path.join(r, file))

    cnt = 0

    for f in files:
        cnt = cnt + 1
        if cnt == x:
            os.chdir(orig)
            return path+'\\'+f[2:]


def exist(path):
    orig = os.getcwd()

    os.chdir(path)
    lst = os.listdir(path)
    os.chdir(orig)

    if not lst:
        return False
    else:
        return True

def eraseContent(path):

    for root, dirs, files in os.walk(path):
        for f in files:
            os.unlink(os.path.join(root, f))
        for d in dirs:
            shutil.rmtree(os.path.join(root, d))

def eraseDirectory(path):
    shutil.rmtree(path)


def convertXlsToXlsx(file):

    if getExtension(file) != "xls":
        return

    f = xlrd.open_workbook(file)

    sheet = f.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols

    new_f = openpyxlWorkbook()
    sheet1 = new_f.get_active_sheet()

    for row in range(0, nrows):
        for col in range(0, ncols):
            sheet1.cell(row=row+1, column=col+1).value = sheet.cell_value(row, col)

    new_f.save(file+'x')

def getExtension(file):

    extension = os.path.splitext(file)[1][1:]
    return extension

def convertAll(path):

    files = []

    for folder, subs, filees in os.walk(path):
        for filename in filees:
            files.append(os.path.abspath(os.path.join(folder, filename)))

    for f in files:
        convertXlsToXlsx(f)

def iter_rows(ws):
    for row in ws.iter_rows():
        yield [cell.value for cell in row]


def addFileToOtherFile(file,other_file,coloanaBarcode,randInceput = 2,randTitluri = 1): #aceasta functie presupune ca ambele au aceeasi extensie (.xls/.xlsx) asa ca trebuie apelata mai intai functia de conversie

    orig = os.getcwd()


    other_file_workbook = load_workbook(other_file)
    other_file_sheet = other_file_workbook.active

    file_workbook = load_workbook(file)
    file_sheet = file_workbook.active

    for row in list(iter_rows(other_file_sheet)):
        file_sheet.append(row)

    file_workbook.save(filename = file)

    #os.remove(other_file)

    os.chdir(orig)




def mergeFiles(pathInitial,pathSalvare,coloanaBarcode,randInceput = 2,randInceputTitluri = 1):

    #delete files after using them

    orig = os.getcwd()

    os.chdir(pathInitial)

    files = []

    #TO DO verifica cu getExtension(file) daca fisierele la care formezi path-ul absolut sunt doar de tipurile xls si xlsx

    for folder, subs, filees in os.walk(pathInitial):
        for filename in filees:
            files.append(os.path.abspath(os.path.join(folder, filename)))

    #print(files) aici in files avem path-urile absolute ale tuturor fisierelor dintr-un directoriu (ar trebuie sa fie toate xls si xlsx, dar trebuie o verificare in plus)

    os.chdir(orig)


#other stuff<------------------------------------->!

#File Management Methods Methods <------------------------------------->!
