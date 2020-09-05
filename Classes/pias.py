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

class Pias(MainApplication):

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

        self.createLabelAtPosition(1,0,"Path Preturi: ",50,20)
        self.createLabelAtPosition(1,1,self.supplierInfo['price_folder_path'])
        self.createLabelAtPosition(1,2,"         ")
        self.createButtonAtPosition(1,3,"Browse")
        self.addNormalCommandToButton(1,self.browsePriceFolderFunction)

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
        modifyJson("Pias","catalogue_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(1,self.master.filename.replace('/','\\'))

    def browsePriceFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['price_folder_path'], title="Select Price Folder")
        if self.master.filename == "":
            return
        modifyJson("Pias","price_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(4,self.master.filename.replace('/','\\'))

    def browseFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['save_path'], title="Select Save Path")
        if self.master.filename == "":
            return
        modifyJson("Pias","save_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(7,self.master.filename.replace('/','\\'))

    def browseUpdateTabelFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['tabel_update_path'], title="Select Update Table Path")
        if self.master.filename == "":
            return
        modifyJson("Pias","tabel_update_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(10,self.master.filename.replace('/','\\'))

    def solve(self):

        self.initJsonStuff()

        error = open(MainApplication.univPath+"\\Resources"+"\\error.txt","w")
        PRICE_ERROR = 696969696969
        BARCODE_ERROR = 797979797979

        folder_catalog = self.supplierInfo['catalogue_folder_path']
        if fileCount(folder_catalog) > 1:
            logError("Too many files in \"Catalog\" folder, please only have one file!")
            return
        elif fileCount(folder_catalog) == 0:
            logError("Folder \"Catalog\" is empty, please place the catalog file inside!")
            return
        file_catalog = getFileXFromPath(folder_catalog, 1)

        folder_price = self.supplierInfo['price_folder_path']
        if fileCount(folder_price) > 1:
            logError("Too many files in \"Price\" folder, please only have one file!")
            return
        elif fileCount(folder_price) == 0:
            logError("Folder \"Price\" is empty, please place the price file inside!")
            return
        file_price = getFileXFromPath(folder_price, 1)

        save_path = self.supplierInfo['save_path']
        start_row = self.supplierInfo['start_row']
        price_start_row = self.supplierInfo['price_start_row']
        barcode_column = self.supplierInfo['barcode_column']
        price_column = self.supplierInfo['price_column']
        pricecode_column = self.supplierInfo['pricecode_column']
        rounded_price_column = self.supplierInfo['rounded_price_column']
        quantity_column = self.supplierInfo['quantity_column']
        noutati_raft_price_column = self.supplierInfo['noutati_raft_price_column']
        raft_price_column = self.supplierInfo['raft_price_column']
        save_name = self.supplierInfo['save_name']

        catalogExt = getExtension(file_catalog)

        deleteNullQuantityFromFile(file_catalog, start_row, quantity_column)

        priceExt = getExtension(file_price)

        piasCatalog = SupplierFile(file_catalog, catalogExt, start_row)
        piasPrices = SupplierFile(file_price, priceExt, price_start_row)

        dictPreturi = piasPrices.getDictionary(pricecode_column,rounded_price_column)

        piasCatalogBarcodes = piasCatalog.getBarcodeDictionary(barcode_column)


        ok = False
        ######################################

        folder_tabel = self.supplierInfo['tabel_update_path']
        file_count = fileCount(folder_tabel)

        if file_count > 1:
            logError("Too many files in \"Tabel\" folder, please only have one file!")
            return

        noutati_start_row = self.supplierInfo['noutati_start_row']
        tabel_barcode_column = self.supplierInfo['tabel_barcode_column']
        tabel_artist_column = self.supplierInfo['tabel_artist_column']
        tabel_title_column = self.supplierInfo['tabel_title_column']
        tabel_suport_column = self.supplierInfo['tabel_suport_column']
        tabel_unit_column = self.supplierInfo['tabel_unit_column']
        tabel_release_date_column = self.supplierInfo['tabel_release_date_column']
        tabel_pricecode_column = self.supplierInfo['tabel_pricecode_column']
        tabel_start_row = self.supplierInfo['tabel_start_row']

        folder_noutati = MainApplication.univPath + "\\Noutati\\Pias"
        file_noutati = getFileXFromPath(folder_noutati, 1)
        noutatiExt = getExtension(file_noutati)
        piasNoutati = SupplierFile(file_noutati, noutatiExt, noutati_start_row)
        noutatiBarcodeDict = piasNoutati.getBarcodeDictionary(tabel_barcode_column)
        startNoutatiBarcodeDict = piasNoutati.getBarcodeDictionary(tabel_barcode_column)

        if file_count == 1:

            ok = True
            file_tabel = getFileXFromPath(folder_tabel, 1)

            tabelExt = getExtension(file_tabel)
            piasTabel = SupplierFile(file_tabel, tabelExt, tabel_start_row)

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
            #raft listare
            try:
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                workbook = excel.Workbooks.Open(os.path.abspath(file_noutati))
                workbook.Save()
                workbook.Close()
                excel.Quit()
            except:
                logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
                return

            deleteBarcodesFromFile(file_tabel, tabel_start_row, tabel_barcode_column, startNoutatiBarcodeDict)

            wb = load_workbook(file_tabel)
            ws = wb.active

            prow = ws.max_row+1
            pcol = ws.max_column+1

            for i in range(tabel_start_row+1, prow):
                correct_price_name = (ws.cell(row = i, column = tabel_pricecode_column).value)
                if ',' in str(correct_price_name):
                    correct_price_name = str(correct_price_name).replace(',','.')
                    correct_price_name = float(correct_price_name)
                if correct_price_name in dictPreturi:
                    price = dictPreturi[correct_price_name]
                else:
                    price = PRICE_ERROR
                if str(price) == None:
                    price = PRICE_ERROR
                else:
                    if not isfloat(str(price)):
                        price = PRICE_ERROR
                print(price)
                if price != PRICE_ERROR:
                    ws.cell(row = i, column = noutati_raft_price_column).value = price

            wb.save(file_tabel)

            try:
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                workbook = excel.Workbooks.Open(os.path.abspath(file_tabel))
                workbook.Save()
                workbook.Close()
                excel.Quit()
            except:
                logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
                return

            if getExtension(file_tabel) == "xls":
                shutil.copy2(file_tabel, save_path+"\\" + "PiasListareNoutati.xls")
                noutati_listare = save_path+"\\" + "PiasListareNoutati.xls"
            elif getExtension(file_tabel) == "xlsx":
                shutil.copy2(file_tabel, save_path+"\\" + "PiasListareNoutati.xlsx")
                noutati_listare = save_path+"\\" + "PiasListareNoutati.xlsx"
            else:
                logError("Failed to copy tabel file")
            #raft listare
            eraseContent(folder_tabel)


        deleteBarcodesFromFile(file_noutati, noutati_start_row, tabel_barcode_column, piasCatalogBarcodes)

        if getExtension(file_catalog) == "xls":
            shutil.copy2(file_catalog, save_path+"\\" + "PiasListare.xls")
            file_listare = save_path+"\\" + "PiasListare.xls"
        elif getExtension(file_catalog) == "xlsx":
            shutil.copy2(file_catalog, save_path+"\\" + "PiasListare.xlsx")
            file_listare = save_path+"\\" + "PiasListare.xlsx"
        else:
            logError("Failed to copy catalog file")

        try:
            qualityConvertXlsToXlsx(file_listare)
        except:
            logError("Conversion did not happen on file_listare")
        if getExtension(file_listare) == 'xls':
            file_listare+='x'
        wb = load_workbook(file_listare)
        ws = wb.active

        prow = ws.max_row+1
        pcol = ws.max_column+1

        for i in range(start_row, prow):
            correct_price_name = (ws.cell(row = i, column = price_column).value)

            if ',' in str(correct_price_name):
                correct_price_name = str(correct_price_name).replace(',','.')
                correct_price_name = float(correct_price_name)
            if correct_price_name in dictPreturi:
                price = dictPreturi[correct_price_name]
            else:
                price = PRICE_ERROR

            if str(price) == None:
                price = PRICE_ERROR
            else:
                if not isfloat(str(price)):
                    price = PRICE_ERROR

            if price != PRICE_ERROR:
                ws.cell(row = i, column = raft_price_column).value = price

        wb.save(file_listare)

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(os.path.abspath(file_listare))
            workbook.Save()
            workbook.Close()
            excel.Quit()
        except:
            logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
            return
        ######################################

        void_workbook = openpyxlWorkbook()
        void_sheet = void_workbook.active
        void_sheet.cell(row = 1,column = 1).value = "Barcode"
        void_sheet.cell(row = 1,column = 6).value = "Price"


        currentRow = 1
        i = start_row-1

        for row in (piasCatalog.data):
            i = i + 1

            barcode = str(row[barcode_column-1])
            catalog_price = row[price_column-1]


            if ',' in str(catalog_price):
                catalog_price = str(catalog_price).replace(',','.')
                catalog_price = float(catalog_price)

            if catalog_price in dictPreturi:
                price = dictPreturi[catalog_price]
            else:
                price = PRICE_ERROR

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
            wb = load_workbook(file_noutati)
            ws = wb.active

            prow = ws.max_row + 1

            for i in range(noutati_start_row + 1, prow):

                barcode = normalizeBarcode(str(ws.cell(row = i, column = tabel_barcode_column).value))
                catalog_price = ws.cell(row = i, column = tabel_pricecode_column).value

                if ',' in str(catalog_price):
                    catalog_price = str(catalog_price).replace(',','.')
                    catalog_price = float(catalog_price)

                if catalog_price in dictPreturi:
                    price = dictPreturi[catalog_price]
                else:
                    price = PRICE_ERROR

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
