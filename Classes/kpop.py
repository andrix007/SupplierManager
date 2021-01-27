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
from selenium import webdriver

class Kpop(MainApplication):

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

        self.createLabelAtPosition(2,0,"Path Download: ",50,20)
        self.createLabelAtPosition(2,1,self.supplierInfo['download_path'])
        self.createLabelAtPosition(2,2,"         ")
        self.createButtonAtPosition(2,3,"Browse")
        self.addNormalCommandToButton(2,self.browseDownloadFunction)

        self.createLabelAtPosition(3,0,"Path Salvare: ",50,20)
        self.createLabelAtPosition(3,1,self.supplierInfo['save_path'])
        self.createLabelAtPosition(3,2,"         ")
        self.createButtonAtPosition(3,3,"Browse")
        self.addNormalCommandToButton(3,self.browseFolderFunction)

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
        modifyJson("Kpop","catalogue_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(1,self.master.filename.replace('/','\\'))


    def browsePriceFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['price_folder_path'], title="Select Price Folder")
        if self.master.filename == "":
            return
        modifyJson("Kpop","price_folder_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(4,self.master.filename.replace('/','\\'))


    def browseFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['save_path'], title="Select Save Path")
        if self.master.filename == "":
            return
        modifyJson("Kpop","save_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(10,self.master.filename.replace('/','\\'))


    def browseDownloadFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['download_path'], title="Select Downloads Folder")
        if self.master.filename == "":
            return
        modifyJson("Kpop","download_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(7,self.master.filename.replace('/','\\'))

    def importFileFromWebsite(self):

        if fileCount(MainApplication.univPath+"\\chromedriver") == 0:
            logError("There are no versions of selenium chromedriver installed!\n To install them go to the official selenium page, download the last\n three versions and place them in the chromedriver directory inside the project.")
            return

        chromedriverPaths = []

        for root, dirs, files in os.walk(MainApplication.univPath+"\\chromedriver"):
            for name in files:
                full_path = os.path.join(root, name)
                chromedriverPaths.append(full_path)

        ok = False

        for i in range(len(chromedriverPaths)):
            try:
                driver = webdriver.Chrome(chromedriverPaths[i])
                ok = True
                print("Version",getVersion(chromedriverPaths[i]),"has executed successfully!")
                break
            except:
                print("Version",getVersion(chromedriverPaths[i]),"doesn't work!")
                continue

        if ok == False:
            logError("The versions of the selenium chromedriver are outdated!\n To install newer ones go to the official selenium page, download the last\n three versions and place them in the chromedriver directory inside the project.")
            return

        try:
            driver.get('http://btb.cnlmusic.kr/btb/bbs/login.php');
        except:
            logError('Failed to connect to website!')
            return
        time.sleep(0.5)

        user=""
        password=""

        try:
            with open(MainApplication.univPath+'\\Resources\\login_data.txt') as f:
                lines = f.readlines()

                user = lines[0].replace('\r','').replace('\n','')
                password = lines[1].replace('\r','').replace('\n','')
        except:
            logError("There is no login_data file in the \\Resources subdirectory! Please place a .txt file called login_data.txt\n with the username on the first line and the password on the second!")
            return

        try:
            driver.find_element_by_id("login_id").send_keys(user)
            driver.find_element_by_id("login_pw").send_keys(password)
            driver.find_element_by_xpath("//*[@id=\"mb_login\"]/div/div/div/section/form/div[3]/input").click()
            time.sleep(0.5)
            driver.find_element_by_xpath("//*[@id=\"sidebar-menu\"]/div/ul/li[2]/a").click()
            time.sleep(0.5)
            driver.find_element_by_xpath("/html/body/div[1]/div/div[3]/div/div[2]/div/div/div[1]/div[2]/a").click()
        except:
            logError("The layout of the website has been changed,\n please consult your lead programmer for fix!")
            return

        while getExtension(newest(self.supplierInfo['download_path'])) != "xls" and  getExtension(newest(self.supplierInfo['download_path'])) != "xlsx" and getName(newest(self.supplierInfo['download_path']))[:10] != 'new_release':
            time.sleep(1)

        driver.quit()

        file = newest(self.supplierInfo['download_path'])

        eraseContent(MainApplication.univPath+"\\temp")

        self.supplierInfo['temp_path'] = MainApplication.univPath+"\\temp\\"+getName(file)
        shutil.move(file, self.supplierInfo['temp_path'])

        try:
            qualityConvertXlsToXlsx(self.supplierInfo['temp_path'])
        except:
            logError('Conversion of new_release_list from xls to xlsx failed!')
            return

        self.supplierInfo['temp_path'] = self.supplierInfo['temp_path']+'x'


    def solve(self):

        self.initJsonStuff()

        self.importFileFromWebsite()

        print("File has been imported successfully!")

        error = open(MainApplication.univPath+"\\Resources"+"\\error.txt","w")
        PRICE_ERROR = 696969696969
        BARCODE_ERROR = 797979797979
        blacklistFormat = ['BOOK', 'KIHNO']

        new_file = self.supplierInfo['temp_path']
        new_start_row = self.supplierInfo['new_start_row']

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
            logError("Folder \"Price\" is empty, please place the catalog file inside!")
            return
        file_price = getFileXFromPath(folder_price, 1)

        save_path = self.supplierInfo['save_path']
        start_row = self.supplierInfo['start_row']
        price_start_row = self.supplierInfo['price_start_row']
        barcode_column = self.supplierInfo['barcode_column']
        new_barcode_column = self.supplierInfo['new_barcode_column']
        format_column = self.supplierInfo['format_column']
        save_name = self.supplierInfo['save_name']
        new_format_column = self.supplierInfo['new_format_column']
        rounded_price_column = self.supplierInfo['rounded_price_column']
        price_file_price_column = self.supplierInfo['price_file_price_column']
        quantity_column = self.supplierInfo['quantity_column']
        new_quantity_column = self.supplierInfo['new_quantity_column']
        artist_column = self.supplierInfo['artist_column']
        new_artist_column = self.supplierInfo['new_artist_column']
        title_column = self.supplierInfo['title_column']
        new_title_column = self.supplierInfo['new_title_column']
        release_date_column = self.supplierInfo['release_date_column']
        new_release_date_column = self.supplierInfo['new_release_date_column']
        catalog_number_column = self.supplierInfo['catalog_number_column']
        new_catalog_number_column = self.supplierInfo['new_catalog_number_column']
        label_column = self.supplierInfo['label_column']
        new_label_column = self.supplierInfo['new_label_column']
        price_column = self.supplierInfo['price_column']
        new_price_column = self.supplierInfo['new_price_column']
        country_of_pressing_column = self.supplierInfo['country_of_pressing_column']
        currency_column = self.supplierInfo['currency_column']
        genre_column = self.supplierInfo['genre_column']
        raft_price_column = self.supplierInfo['raft_price_column']
        price_file_void_column = self.supplierInfo['price_file_void_column']
        void_column = self.supplierInfo['void_column']

        catalogExt = getExtension(file_catalog)
        newExt = getExtension(new_file)
        priceExt = getExtension(file_price)

        kpopCatalog = SupplierFile(file_catalog, catalogExt, start_row)
        kpopPrices = SupplierFile(file_price, priceExt, price_start_row)
        new_release_catalog = SupplierFile(new_file, newExt, new_start_row)

        catalogBarcodeDict = kpopCatalog.getBarcodeDictionary(barcode_column)


        #Prima reforma in care pun numai elementele relevante in new_release

        pureNewRelease = openpyxlWorkbook()
        pureNewSheet = pureNewRelease.active

        wb = load_workbook(new_file)
        ws = wb.active

        prow = ws.max_row+1
        pcol = ws.max_column+1
        cnt = new_start_row

        for i in range(new_start_row, prow):
            if ws.cell(row = i, column = barcode_column).value not in catalogBarcodeDict and ws.cell(row = i, column = new_format_column).value not in blacklistFormat:
                for j in range(1, pcol):
                    pureNewSheet.cell(row = cnt, column = j).value = ws.cell(row = i, column = j).value
                cnt = cnt + 1

        pureNewRelease.save(new_file)

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(os.path.abspath(new_file))
            workbook.Save()
            workbook.Close()
            excel.Quit()
        except:
            logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
            return

        #!Prima reforma in care pun numai elementele relevante in new_release

        priceDict = kpopPrices.getDictionary(price_file_price_column, rounded_price_column)
        voidPriceDict = kpopPrices.getDictionary(price_file_price_column, price_file_void_column)

        #Fac new release-ul sa arate cum dtrebuie

        pureNewRelease = openpyxlWorkbook()
        pureNewSheet = pureNewRelease.active

        try:
            wb = load_workbook(new_file)
        except:
            logError("new_release list is either missing or open in another program!")
            return

        ws = wb.active

        prow = ws.max_row+1
        pcol = ws.max_column+1
        cnt = new_start_row

        for i in range(new_start_row, prow):

                pureNewSheet.cell(row = cnt, column = barcode_column).value = str(ws.cell(row = i, column = new_barcode_column).value)
                pureNewSheet.cell(row = cnt, column = format_column).value = ws.cell(row = i, column = new_format_column).value
                pureNewSheet.cell(row = cnt, column = quantity_column).value = ws.cell(row = i, column = new_quantity_column).value
                pureNewSheet.cell(row = cnt, column = artist_column).value = ws.cell(row = i, column = new_artist_column).value
                pureNewSheet.cell(row = cnt, column = title_column).value = ws.cell(row = i, column = new_title_column).value
                pureNewSheet.cell(row = cnt, column = release_date_column).value = ws.cell(row = i, column = new_release_date_column).value
                pureNewSheet.cell(row = cnt, column = catalog_number_column).value = ws.cell(row = i, column = new_catalog_number_column).value
                pureNewSheet.cell(row = cnt, column = label_column).value = ws.cell(row = i, column = new_label_column).value
                _price = getCorrectPrice(ws.cell(row = i, column = new_price_column).value)
                pureNewSheet.cell(row = cnt, column = price_column).value = _price
                pureNewSheet.cell(row = cnt, column = country_of_pressing_column).value = "KOREA"
                pureNewSheet.cell(row = cnt, column = currency_column).value = "KRW"
                pureNewSheet.cell(row = cnt, column = genre_column).value = "K-POP"
                pureNewSheet.cell(row = cnt, column = raft_price_column).value = priceDict[_price]
                pureNewSheet.cell(row = cnt, column = void_column).value = voidPriceDict[_price]

                cnt = cnt + 1

        pureNewRelease.save(new_file)

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(os.path.abspath(new_file))
            workbook.Save()
            workbook.Close()
            excel.Quit()
        except:
            logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
            return
        #de scos kihno si book din catalog

        pureCatalog = openpyxlWorkbook()
        pureNewSheet = pureCatalog.active

        try:
            wb = load_workbook(file_catalog)
        except:
            logError("catalog file is either missing or open in another program!")
            return
        ws = wb.active

        prow = ws.max_row+1
        pcol = ws.max_column+1
        cnt = start_row

        for j in range(1, pcol):
            pureNewSheet.cell(row = start_row-1, column = j).value = ws.cell(start_row-1, column = j).value

        for i in range(start_row, prow):
            f = getCorrectFormat(ws.cell(row = i, column = format_column).value)

            if f != 'KIHNO' and f != 'BOOK':
                for j in range(1, pcol):
                    pureNewSheet.cell(row = cnt, column = j).value = ws.cell(row = i, column = j).value

                cnt = cnt + 1

        pureCatalog.save(file_catalog)

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(os.path.abspath(file_catalog))
            workbook.Save()
            workbook.Close()
            excel.Quit()
        except:
            logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
            return



        #de adaugat new release la catalog
        try:
            catWB = load_workbook(file_catalog)
        except:
            logError("catalog file is either missing or open in another program!")
            return
        catWS = catWB.active

        try:
            wb = load_workbook(new_file)
        except:
            logError("new_release list is either missing or open in another program!")
            return

        ws = wb.active

        crow = catWS.max_row+1
        ccol = catWS.max_column+1

        prow = ws.max_row+1
        pcol = ws.max_column+1

        cnt = new_start_row

        for i in range(crow, crow+prow):
            for j in range(1, pcol):
                catWS.cell(row = i, column = j).value = ws.cell(row = i-crow+new_start_row, column = j).value

        catWB.save(file_catalog)

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            workbook = excel.Workbooks.Open(os.path.abspath(file_catalog))
            workbook.Save()
            workbook.Close()
            excel.Quit()
        except:
            logError("Problem with Microsoft Excel!\n Also, if any file that might be used by the program is open,\n please close it and try again!")
            return

        #PRELUCRARE

        catalogExt = getExtension(file_catalog)
        kpopCatalog = SupplierFile(file_catalog, catalogExt, start_row)

        void_workbook = openpyxlWorkbook()
        void_sheet = void_workbook.active
        void_sheet.cell(row = 1,column = 1).value = "Barcode"
        void_sheet.cell(row = 1,column = 2).value = "Price"

        currentRow = 1
        i = start_row-1

        for row in (kpopCatalog.data):
            i = i + 1

            barcode = str(row[barcode_column-1])
            price = row[void_column-1]

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
        void_workbook.save(os.path.join(save_path, save_name))

        print("Void file has been created successfully!")

        self.master.destroy()
        logText("Code has executed successfully!")


