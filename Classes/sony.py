if __name__ == "__main__":
    from MainApplication import *
else:
    from Classes.MainApplication import *
if __name__ == "__main__":
    from functions import *
else:
    from Classes.functions import *

import os
import json

class Sony(MainApplication):

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

        self.supplierInfo = {}

        with open(MainApplication.jsonFilePath) as f:
            data = json.load(f)

        for state in data["suppliers"]:
            if state.get('title') == title:
                self.supplierInfo = state
                break

        #print(self.supplierInfo)

    def solve(self):

        print(getFileXFromPath('C:\\Users\\Andrei Bancila\\Desktop\\Furnizori\\Plaseaza_fisierele_aici\\Catalog',1))
        convertXlsToXlsx(getFileXFromPath('C:\\Users\\Andrei Bancila\\Desktop\\Furnizori\\Plaseaza_fisierele_aici\\Catalog',1))
        print(getExtension(getFileXFromPath('C:\\Users\\Andrei Bancila\\Desktop\\Furnizori\\Plaseaza_fisierele_aici\\Catalog',1)))

        #mergeFiles('C:\\Users\\Andrei Bancila\\Desktop\\Furnizori\\Plaseaza_fisierele_aici\\Catalog','C:\\Users\\Andrei Bancila\\Desktop\\Furnizori\\Plaseaza_fisierele_aici\\Catalog',1,1)

        addFileToOtherFile(getFileXFromPath('C:\\Users\\Andrei Bancila\\Desktop\\Furnizori\\Plaseaza_fisierele_aici\\Catalog',1),getFileXFromPath('C:\\Users\\Andrei Bancila\\Desktop\\Furnizori\\Plaseaza_fisierele_aici\\Catalog',2),4)
        self.master.destroy()
