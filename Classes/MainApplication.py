import tkinter as tk
from tkinter import *

class MainApplication:



    #Class Variables <------------------------------------->



    univBackColor = "#012e20"
    univForColor = "white"


    #Class Variables <------------------------------------->!



    def __init__(self, master,title, width, height, icon=None, color=None):
        self.master = master
        self.master.geometry(str(width)+"x"+str(height))
        self.master.title(title)
        self.master.iconbitmap(icon)
        self.master.config(bg = color)
        self.universalBackgroundColor = MainApplication.univBackColor
        self.universalForegroundColor = MainApplication.univForColor

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
        drop['menu'].config(bg=self.universalBackgroundColor,fg=self.universalForegroundColor,activebackground = self.universalBackgroundColor, activeforeground = self.universalForegroundColor)

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



    def getTextFromEntry(self,index=0):
        text = self.entries[index].get()
        self.entries[index].delete(0)
        return text



    def addNormalCommandToButton(self,index,command):
        self.buttons[index].config(command = command)



    def addLambdaCommandToButton(self,index,lambda_command,*args):
        self.buttons[index].config(command = lambda : lambda_command(*args))



    def showWorks(self,text):
        print(text.upper())


    def showSelection(self,index):
        text = self.ddmenus_clicked[index].get()
        print(text)

    def getSelection(self,index):
        text = self.ddmenus_clicked[index].get()
        return text

    #Widget Management Methods <------------------------------------->!



    #Just Stuff Methods <------------------------------------->



    #Just Stuff Methods <------------------------------------->



    #Class Methods <------------------------------------->



    @classmethod
    def changeUnivBackColor(cls,color):
        cls.univBackColor = color



    @classmethod
    def changeUnivForColor(cls,color):
        cls.univForColor = color



    #Class Methods <------------------------------------->!

