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
class SCR(MainApplication):
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
        self.createLabelAtPosition(2,0,"Path Update: ",50,20)
        self.createLabelAtPosition(2,1,self.supplierInfo['update_path'])
        self.createLabelAtPosition(2,2,"         ")
        self.createButtonAtPosition(2,3,"Browse")
        self.addNormalCommandToButton(2,self.browseUpdateFolderFunction)
        self.createLabelAtPosition(3,0,"Path Salvare: ",50,20)
        self.createLabelAtPosition(3,1,self.supplierInfo['save_path'])
        self.createLabelAtPosition(3,2,"         ")
        self.createButtonAtPosition(3,3,"Browse")
        self.addNormalCommandToButton(3,self.browseFolderFunction)
        self.createLabelAtPosition(4,0,"                    ")
        self.createButtonAtPosition(4,1,"Solve")
        self.addNormalCommandToButton(4,self.solve)
        self.createLabelAtPosition(4,2,"                    ")

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
        modifyJson("Speakers Corner","catalogue_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(1,self.master.filename.replace('/','\\'))


    def browsePriceFolderFunction(self):
        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['price_folder_path'], title="Select Price Folder")
        if self.master.filename == "":
            return
        modifyJson("Speakers Corner","price_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(4,self.master.filename.replace('/','\\'))

    def browseUpdateFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['update_path'], title="Select Save Path")
        if self.master.filename == "":
            return
        modifyJson("Speakers Corner","update_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(7,self.master.filename.replace('/','\\'))

    def browseFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['save_path'], title="Select Save Path")
        if self.master.filename == "":
            return
        modifyJson("Speakers Corner","save_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(10,self.master.filename.replace('/','\\'))

    def solve(self):
        self.initJsonStuff()
        error = open(MainApplication.univPath+"\\Resources"+"\\error.txt","w")
        PRICE_ERROR = 696969696969
        BARCODE_ERROR = 797979797979
        catalog_folder = self.supplierInfo['catalogue_folder_path']
        if fileCount(catalog_folder) > 1:
            logError("Too many files in \"Catalog\" folder, please only have one file!")
            return
        elif fileCount(catalog_folder) == 0:
            logError("Folder \"Catalog\" is empty, please place the catalog file inside!")
            return
        file_catalog = getFileXFromPath(catalog_folder, 1)
        price_folder = self.supplierInfo['price_folder_path']
        if fileCount(price_folder) > 1:
            logError("Too many files in \"Price\" folder, please only have one file!")
            return
        elif fileCount(price_folder) == 0:
            logError("Folder \"Price\" is empty, please place the price file inside!")
            return
        file_price = getFileXFromPath(price_folder, 1)

        update_folder = self.supplierInfo['update_path']

        if fileCount(update_folder) > 3:
            logError("Too many files in \"Update\" folder, please only have the 3 files: New, Release and Repress!")
            return
        elif fileCount(update_folder) < 3 and fileCount(update_folder) > 0:
            logError("Folder \"Update\" is missing some files, please place all the 3 following files inside: New, Release and Repress!")
            return
        elif fileCount(update_folder) == 0:
            logError("Folder \"Update\" is empty, please place the 3 files (New, Release and Repress) inside!")
            return

        file_new = getFileXFromPath(update_folder, 1)
        file_release = getFileXFromPath(update_folder, 2)
        file_repress = getFileXFromPath(update_folder, 3)
        save_path = self.supplierInfo['save_path']
        start_row = self.supplierInfo['start_row']
        price_start_row = self.supplierInfo['price_start_row']
        barcode_column = self.supplierInfo['barcode_column']
        price_column = self.supplierInfo['price_column']
        pricecode_column = self.supplierInfo['pricecode_column']
        rounded_price_column = self.supplierInfo['rounded_price_column']
        save_name = self.supplierInfo['save_name']
        your_reference_column = self.supplierInfo['your_reference_column']
        raft_listare_price_column = self.supplierInfo['raft_listare_price_column']
        noutati_raft_price_column = self.supplierInfo['noutati_raft_price_column']
        catalogExt = getExtension(file_catalog)
        priceExt = getExtension(file_price)
        newExt = getExtension(file_new)
        releaseExt = getExtension(file_release)
        repressExt = getExtension(file_repress)

        scrCatalog = SupplierFile(file_catalog,catalogExt,start_row)
        scrPrices = SupplierFile(file_price,priceExt,price_start_row)
        scrNew = SupplierFile(file_new,newExt,start_row)
        scrRelease = SupplierFile(file_release,releaseExt,start_row)
        scrRepress = SupplierFile(file_repress,repressExt,start_row)

        mergeContainerFile = SupplierFile()

        scrNew.printColumn(your_reference_column)
        scrRelease.printColumn(your_reference_column)
        scrRepress.printColumn(your_reference_column)

        if not scrNew.isColumnEmpty(your_reference_column):
            mergeContainerFile.addOtherSupplierFile(scrNew)
        else:
            mergeContainerFile.addOtherSupplierFile(scrNew, your_reference_column-1)
        if not scrRelease.isColumnEmpty(your_reference_column):
            mergeContainerFile.addOtherSupplierFile(scrRelease)
        else:
            mergeContainerFile.addOtherSupplierFile(scrRelease, your_reference_column-1)
        if not scrRepress.isColumnEmpty(your_reference_column):
            mergeContainerFile.addOtherSupplierFile(scrRepress)
        else:
            mergeContainerFile.addOtherSupplierFile(scrRepress, your_reference_column-1)

        dictProductCode = mergeContainerFile.getDictionary(barcode_column, price_column)

        i = 0

        for row in scrCatalog.data:

            barcode = str(normalizeBarcode(row[barcode_column-1]))
            pricecode = row[price_column-1]

            if barcode in dictProductCode:
                if dictProductCode[barcode] != pricecode:
                    scrCatalog.data[i][price_column-1] = dictProductCode[barcode]

            i = i + 1


        dictBarcode = scrCatalog.getBarcodeDictionary(barcode_column)
        dictPreturi = scrPrices.getDictionary(pricecode_column, rounded_price_column)

        file_noutati = save_path+"\\SpeakersCornerNoutati.xlsx"
        scrCatalog.addOtherSupplierFileInRegardToDictionarySCR(mergeContainerFile, dictBarcode, barcode_column, file_noutati)
        scrCatalog.addContentToWorkbook(file_catalog, start_row)
        #raft file catalog

        if getExtension(file_catalog) == "xls":
            shutil.copy2(file_catalog, save_path+"\\" + "SpeakersCornerListare.xls")
            file_listare = save_path+"\\" + "SpeakersCornerListare.xls"
        elif getExtension(file_catalog) == "xlsx":
            shutil.copy2(file_catalog, save_path+"\\" + "SpeakersCornerListare.xlsx")
            file_listare = save_path+"\\" + "SpeakersCornerListare.xlsx"
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
            catalog_price = ws.cell(row = i, column = price_column).value
            if catalog_price in dictPreturi:
                price = dictPreturi[catalog_price]
            else:
                price = PRICE_ERROR
            if price != PRICE_ERROR:
                ws.cell(row = i, column = raft_listare_price_column).value = price

        wb.save(file_listare)
        #raft file catalog
        #raft file noutati
        wb = load_workbook(file_noutati)
        ws = wb.active

        prow = ws.max_row+1
        pcol = ws.max_column+1

        for i in range(start_row, prow):
            catalog_price = ws.cell(row = i, column = price_column).value
            if catalog_price in dictPreturi:
                price = dictPreturi[catalog_price]
            else:
                price = PRICE_ERROR
            if price != PRICE_ERROR:
                ws.cell(row = i, column = noutati_raft_price_column).value = price

        wb.save(file_noutati)
        #raft file noutati


        void_workbook = openpyxlWorkbook()
        void_sheet = void_workbook.active
        void_sheet.cell(row = 1,column = 1).value = "Barcode"
        void_sheet.cell(row = 1,column = 2).value = "Price"


        currentRow = 1
        i = start_row-1
        dictBarcode = {}
        for row in (scrCatalog.data):
            i = i + 1
            barcode = str(row[barcode_column-1])
            catalog_price = row[price_column-1]
            if catalog_price in dictPreturi:
                price = dictPreturi[catalog_price]
            else:
                price = PRICE_ERROR
            barcode = normalizeBarcode(barcode)
            if barcode == None:
                barcode = BARCODE_ERROR
            else:
                if not barcode.isdigit():
                    print("At line",i,"there is a problem with the barcode",barcode)
                    print(barcode.isdigit())
                    barcode = BARCODE_ERROR
                else:
                    if barcode in dictBarcode:
                        print("At line",i,"barcode",barcode,"is already in dictionary")
                    else:
                        dictBarcode.update({barcode:"1"})
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
