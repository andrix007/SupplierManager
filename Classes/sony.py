if __name__ == "__main__":
    from MainApplication import *
else:
    from Classes.MainApplication import *
if __name__ == "__main__":
    from functions import *
else:
    from Classes.functions import *
if __name__ == "__main__":
    from SupplierFile import *
else:
    from Classes.SupplierFile import *

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
        self.title = title

        with open(MainApplication.jsonFilePath) as f:
            data = json.load(f)

        for state in data["suppliers"]:
            if state.get('title') == title:
                self.supplierInfo = state
                break

        self.createLabelAtPosition(0,0,"Path Catalog: ",50,20)
        self.createLabelAtPosition(0,1,self.supplierInfo['catalogue_path'])
        self.createLabelAtPosition(0,2,"         ")
        self.createButtonAtPosition(0,3,"Browse")
        self.addNormalCommandToButton(0,self.browseFunction)

        self.createLabelAtPosition(1,0,"Path Salvare: ",50,20)
        self.createLabelAtPosition(1,1,self.supplierInfo['save_path'])
        self.createLabelAtPosition(1,2,"         ")
        self.createButtonAtPosition(1,3,"Browse")
        self.addNormalCommandToButton(1,self.browseFolderFunction)

        self.createLabelAtPosition(2,0,"                    ")
        self.createButtonAtPosition(2,1,"Solve")
        self.addNormalCommandToButton(2,self.solve)
        self.createLabelAtPosition(2,2,"                    ")
        #print(self.supplierInfo)

    def initJsonStuff(self):
        self.supplierInfo = {}

        with open(MainApplication.jsonFilePath) as f:
            data = json.load(f)

        for state in data["suppliers"]:
            if state.get('title') == self.title:
                self.supplierInfo = state
                break

    def browseFunction(self):

        self.master.filename  = filedialog.askopenfilename(initialdir=self.supplierInfo['catalogue_path'].split('\\')[:-1], title="Select File", filetypes = (("all files","*.*"),("xlsx files","*.xlsx"),("xls files","*.xls")))
        if self.master.filename == "":
            return
        modifyJson("Sony","catalogue_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(1,self.master.filename.replace('/','\\'))

    def browseFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['save_path'], title="Select File")
        if self.master.filename == "":
            return
        modifyJson("Sony","save_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(4,self.master.filename.replace('/','\\'))

    def solve(self):
        self.initJsonStuff()

        error = open(MainApplication.univPath+"\\Resources"+"\\error.txt","w")

        file = self.supplierInfo['catalogue_path']

        save_path = self.supplierInfo['save_path']
        start_row = self.supplierInfo['start_row']
        save_name = self.supplierInfo['save_name']

        extension = getExtension(file)

        sonyCatalog = SupplierFile(file,extension,start_row)

        barcode_column = self.supplierInfo['barcode_column']
        price_column = self.supplierInfo['price_column']

        void_workbook = openpyxlWorkbook()
        void_sheet = void_workbook.active
        void_sheet.cell(row = 1,column = 1).value = "Barcode"
        void_sheet.cell(row = 1,column = 6).value = "Price"

        currentRow = 1
        i = start_row-1

        for row in (sonyCatalog.data):
            i = i + 1

            barcode = str(row[barcode_column-1])
            price = row[price_column-1]

            if barcode.isdigit() and isfloat(str(price)):

                currentRow = currentRow + 1

                barcode = barcode.zfill(13)
                newPrice = price/1.19*0.6

                void_sheet.cell(row = currentRow,column = 1).value = barcode
                void_sheet.cell(row = currentRow,column = 6).value = round(newPrice,2)

            else:

                errorText = "Line " + str(i) + ":   " + barcode + " " + str(price) + "\n"
                error.write(errorText)

        error.close()
        void_workbook.save(save_path+"\\" + save_name)

        self.master.destroy()
