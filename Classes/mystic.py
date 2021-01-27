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

class Mystic(MainApplication):

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
        self.createLabelAtPosition(0,1,self.supplierInfo['catalogue_folder_path'])
        self.createLabelAtPosition(0,2,"         ")
        self.createButtonAtPosition(0,3,"Browse")
        self.addNormalCommandToButton(0,self.browseCatalogFolderFunction)

        self.createLabelAtPosition(1,0,"Path Formula: ",50,20)
        self.createLabelAtPosition(1,1,self.supplierInfo['formula_folder_path'])
        self.createLabelAtPosition(1,2,"         ")
        self.createButtonAtPosition(1,3,"Browse")
        self.addNormalCommandToButton(1,self.browseFormulaFolderFunction)

        self.createLabelAtPosition(2,0,"Path Salvare: ",50,20)
        self.createLabelAtPosition(2,1,self.supplierInfo['save_path'])
        self.createLabelAtPosition(2,2,"         ")
        self.createButtonAtPosition(2,3,"Browse")
        self.addNormalCommandToButton(2,self.browseFolderFunction)

        self.createLabelAtPosition(3,0,"                    ")
        self.createButtonAtPosition(3,1,"Solve")
        self.addNormalCommandToButton(3,self.solve)
        self.createLabelAtPosition(3,2,"                    ")
        #print(self.supplierInfo)

    def initJsonStuff(self):
        self.supplierInfo = {}

        with open(MainApplication.jsonFilePath) as f:
            data = json.load(f)

        for state in data["suppliers"]:
            if state.get('title') == self.title:
                self.supplierInfo = state
                break

    def browseCatalogFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['catalogue_folder_path'], title="Select Catalog Folder")
        if self.master.filename == "":
            return
        modifyJson("Mystic","catalogue_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(1,self.master.filename.replace('/','\\'))

    def browseFormulaFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['formula_folder_path'], title="Select Formula Folder")
        if self.master.filename == "":
            return
        modifyJson("Mystic","formula_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(4,self.master.filename.replace('/','\\'))

    def browseFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['save_path'], title="Select File")
        if self.master.filename == "":
            return
        modifyJson("Mystic","save_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(7,self.master.filename.replace('/','\\'))

    def solve(self):
        self.initJsonStuff()

        error = open(MainApplication.univPath+"\\Resources"+"\\error.txt","w")
        PRICE_ERROR = 696969696969
        BARCODE_ERROR = 797979797979
        blacklisetRecords = ["AFM","PIAS","SNAPPER"]

        folder_catalog = self.supplierInfo['catalogue_folder_path']
        if fileCount(folder_catalog) > 1:
            logError("Too many files in \"Catalog\" folder, please only have one file!")
            return
        elif fileCount(folder_catalog) == 0:
            logError("Folder \"Catalog\" is empty, please place the catalog file inside!")
            return
        file_catalog = getFileXFromPath(folder_catalog, 1)

        folder_formula = self.supplierInfo['formula_folder_path']
        if fileCount(folder_formula) > 1:
            logError("Too many files in \"Formula\" folder, please only have one file!")
            return
        elif fileCount(folder_formula) == 0:
            logError("Folder \"Formula\" is empty, please place the formula file inside!\n (Formula file is the file from which the formula is extracted)")
            return
        file_formula = getFileXFromPath(folder_formula, 1)

        save_path = self.supplierInfo['save_path']
        start_row = self.supplierInfo['start_row']
        records_column = self.supplierInfo['records_column']
        formula_column = self.supplierInfo['formula_column']
        formula_file_start_row = self.supplierInfo['formula_file_start_row']
        barcode_column = self.supplierInfo['barcode_column']
        price_column = self.supplierInfo['price_column']
        save_name = self.supplierInfo['save_name']

        catalogExt = getExtension(file_catalog)
        formulaExt = getExtension(file_formula)

        mysticFormula = SupplierFile(file_formula, formulaExt, formula_file_start_row)

        mysticFormula.formula = mysticFormula.getFormula(formula_file_start_row, formula_column)

        pureCatalog = openpyxlWorkbook()
        pureCatalogSheet = pureCatalog.active

        catalogWB = load_workbook(file_catalog)
        catalogWS = catalogWB.active

        prow = catalogWS.max_row+1
        pcol = catalogWS.max_column+1
        cnt = start_row

        for i in range(start_row, prow):
            if catalogWS.cell(row = i, column = records_column).value not in blacklisetRecords:
                for j in range(1, pcol):
                    pureCatalogSheet.cell(row = cnt, column = j).value = catalogWS.cell(row = i, column = j).value
                pureCatalogSheet.cell(row = cnt, column = formula_column).value = getFormulaForRowX(mysticFormula.formula, 2, cnt)
                cnt = cnt + 1

        pureCatalog.save(file_catalog)

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        workbook = excel.Workbooks.Open(os.path.abspath(file_catalog))
        workbook.Save()
        workbook.Close()
        excel.Quit()


        void_workbook = openpyxlWorkbook()
        void_sheet = void_workbook.active
        void_sheet.cell(row = 1,column = 1).value = "Barcode"
        void_sheet.cell(row = 1,column = 2).value = "Price"


        m_catalog = SupplierFile(file_catalog, catalogExt, start_row)

        currentRow = 1
        i = start_row-1

        for row in m_catalog.data:
            i = i + 1

            barcode = str(row[barcode_column-1])
            price = row[price_column-1]

            barcode = normalizeBarcode(barcode)

            if barcode == None:
                barcode = BARCODE_ERROR
            else:
                if not barcode.isdigit():
                    barcode = BARCODE_ERROR

            if str(price) == None:
                price = PRICE_ERROR
            else:
                if not isfloat(str(price)):
                    price = PRICE_ERROR

            if barcode != BARCODE_ERROR and price != PRICE_ERROR:

                currentRow = currentRow + 1

                barcode = barcode.zfill(13)

                void_sheet.cell(row = currentRow,column = 1).value = barcode
                void_sheet.cell(row = currentRow,column = 2).value = round(price,2)

            else:

                errorText = "Line " + str(i) + ":   "
                if barcode == BARCODE_ERROR:
                    errorText = errorText + "BARCODE_ERROR "
                else:
                    errorText = errorText + barcode + " "

                if price == PRICE_ERROR:
                    errorText = errorText + "PRICE_ERROR "
                else:
                    errorText = errorText + str(price)

                errorText = errorText + "\n"

                error.write(errorText)


        error.close()
        void_workbook.save(save_path+"\\" + save_name)

        self.master.destroy()
        logText("Code has executed successfully!")

