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

        self.createLabelAtPosition(3,0,"Tabel Update Path: ",50,20)
        self.createLabelAtPosition(3,1,self.supplierInfo['tabel_update_path'])
        self.createLabelAtPosition(3,2,"         ")
        self.createButtonAtPosition(3,3,"Browse",)
        self.addNormalCommandToButton(3,self.browseUpdateTabelFunction)

        self.createLabelAtPosition(4,0,"                    ")
        self.createButtonAtPosition(4,1,"Solve")
        self.addNormalCommandToButton(4,self.solve)
        self.createLabelAtPosition(4,2,"                    ")
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

    def browseUpdateTabelFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['tabel_update_path'], title="Select Update Table Path")
        if self.master.filename == "":
            return
        modifyJson("Mystic","tabel_update_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(10,self.master.filename.replace('/','\\'))

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

        mysticCatalog = SupplierFile(file_catalog, catalogExt, start_row)
        mysticCatalogBarcodes = mysticCatalog.getBarcodeDictionary(barcode_column)

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


        ok = False
        ######################################

        folder_tabel = self.supplierInfo['tabel_update_path']
        noutati_start_row = self.supplierInfo['noutati_start_row']
        noutati_formula_column = self.supplierInfo['noutati_formula_column']
        tabel_barcode_column = self.supplierInfo['tabel_barcode_column']
        tabel_artist_column = self.supplierInfo['tabel_artist_column']
        tabel_title_column = self.supplierInfo['tabel_title_column']
        tabel_suport_column = self.supplierInfo['tabel_suport_column']
        tabel_unit_column = self.supplierInfo['tabel_unit_column']
        tabel_release_date_column = self.supplierInfo['tabel_release_date_column']
        tabel_pricecode_column = self.supplierInfo['tabel_pricecode_column']
        tabel_start_row = self.supplierInfo['tabel_start_row']

        folder_noutati = MainApplication.univPath + "\\Noutati\\Mystic"
        file_noutati = getFileXFromPath(folder_noutati, 1)
        noutatiExt = getExtension(file_noutati)
        mysticNoutati = SupplierFile(file_noutati, noutatiExt, noutati_start_row)
        noutatiBarcodeDict = mysticNoutati.getBarcodeDictionary(tabel_barcode_column)

        file_count = fileCount(folder_tabel)

        if file_count > 1:
            logError("Too many files in \"Tabel\" folder, please only have one file!")
            return

        if file_count == 1:

            ok = True
            file_tabel = getFileXFromPath(folder_tabel, 1)


            tabelExt = getExtension(file_tabel)
            mysticTabel = SupplierFile(file_tabel, tabelExt, tabel_start_row)

            try:
                tabel_wb = load_workbook(file_tabel)
            except:
                logError("catalog file is either missing or open in another program!")
                return

            tabel_ws = tabel_wb.active

            tabel_prow = tabel_ws.max_row + 1
            tabel_pcol = tabel_ws.max_column + 1


            try:
                noutati_wb = load_workbook(file_noutati)
            except:
                logError("catalog file is either missing or open in another program!")
                return

            noutati_ws = noutati_wb.active

            noutati_prow  = noutati_ws.max_row + 1
            noutati_pcol = noutati_ws.max_column + 1

            cnt = noutati_prow

            for i in range(tabel_start_row+1, tabel_prow):

                tabel_barcode = normalizeBarcode(str(tabel_ws.cell(row = i, column = tabel_barcode_column).value))
                if tabel_barcode not in noutatiBarcodeDict:
                    for j in range(1, tabel_pcol):
                        noutati_ws.cell(row = cnt, column = j).value = tabel_ws.cell(row = i, column = j).value
                    cnt = cnt + 1
                    noutatiBarcodeDict.update({tabel_barcode : "1"})

            noutati_wb.save(file_noutati)

            try:
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                workbook = excel.Workbooks.Open(os.path.abspath(file_noutati))
                workbook.Save()
                workbook.Close()
                excel.Quit()
            except:
                logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
                return

            pureCatalog = openpyxlWorkbook()
            pureCatalogSheet = pureCatalog.active

            wb = load_workbook(file_noutati)
            ws = wb.active

            prow = ws.max_row+1
            pcol = ws.max_column+1
            cnt = noutati_start_row

            barcodeD = {}

            for i in range(noutati_start_row, prow):
                if i == noutati_start_row:
                    for j in range(1, pcol):
                        pureCatalogSheet.cell(row = cnt, column = j).value = ws.cell(row = i, column = j).value
                    cnt = cnt + 1
                    continue
                barcode = normalizeBarcode(str(ws.cell(row = i, column = tabel_barcode_column).value))
                if barcode not in barcodeD:
                    for j in range(1, pcol):
                        pureCatalogSheet.cell(row = cnt, column = j).value = ws.cell(row = i, column = j).value
                    pureCatalogSheet.cell(row = cnt, column = noutati_formula_column).value = getMysticFormulaForRowX(mysticFormula.formula, 2, cnt)
                    cnt = cnt + 1
                    barcodeD.update({barcode:"1"})

            pureCatalog.save(file_noutati)

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(os.path.abspath(file_noutati))
            workbook.Save()
            workbook.Close()
            excel.Quit()

            #aici sterg din lista cu noutati ce e din pias catalog
            #aici sterg din lista cu noutati ce e din pias catalog

            eraseContent(folder_tabel)

        deleteBarcodesFromFile(file_noutati, noutati_start_row, tabel_barcode_column, mysticCatalogBarcodes)


        void_workbook = openpyxlWorkbook()
        void_sheet = void_workbook.active
        void_sheet.cell(row = 1,column = 1).value = "Barcode"
        void_sheet.cell(row = 1,column = 6).value = "Price"


        m_catalog = SupplierFile(file_catalog, catalogExt, start_row)

        currentRow = 1
        i = start_row-1

        for row in m_catalog.data:
            i = i + 1

            barcode = str(row[barcode_column-1])
            price = row[price_column-1]

            barcode = normalizeBarcode(barcode)

            if barcode == None or barcode == "0000000000000":
                barcode = BARCODE_ERROR
            else:
                if not barcode.isdigit():
                    barcode = BARCODE_ERROR

            if str(price) == None or str(price) == '0' or str(price) == '0.0':
                price = PRICE_ERROR
            else:
                if not isfloat(str(price)):
                    price = PRICE_ERROR

            if barcode != BARCODE_ERROR and price != PRICE_ERROR:

                currentRow = currentRow + 1

                barcode = barcode.zfill(13)

                void_sheet.cell(row = currentRow,column = 1).value = barcode
                void_sheet.cell(row = currentRow,column = 6).value = round(price,2)

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

        if True == True:
            #return
            #aici bag in void tot ce erasi in lista cu noutati
            m_noutati = SupplierFile(file_noutati, noutatiExt, noutati_start_row)

            i = noutati_start_row-1

            for row in m_noutati.data:
                i = i + 1

                try:
                    barcode = str(row[tabel_barcode_column-1])
                except:
                    barcode = BARCODE_ERROR
                try:
                    price = row[noutati_formula_column-1]
                except:
                    price = PRICE_ERROR

                barcode = normalizeBarcode(barcode)

                if barcode == None or barcode == "0000000000000":
                    barcode = BARCODE_ERROR
                else:
                    if not barcode.isdigit():
                        barcode = BARCODE_ERROR

                if str(price) == None or str(price) == '0' or str(price) == '0.0':
                    price = PRICE_ERROR
                else:
                    if not isfloat(str(price)):
                        price = PRICE_ERROR

                if barcode != BARCODE_ERROR and price != PRICE_ERROR:

                    currentRow = currentRow + 1

                    barcode = barcode.zfill(13)

                    void_sheet.cell(row = currentRow,column = 1).value = barcode
                    void_sheet.cell(row = currentRow,column = 6).value = round(price,2)

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

