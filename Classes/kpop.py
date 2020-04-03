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
import time
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
        self.createLabelAtPosition(0,1,self.supplierInfo['catalogue_path'])
        self.createLabelAtPosition(0,2,"         ")
        self.createButtonAtPosition(0,3,"Browse")
        self.addNormalCommandToButton(0,self.browseCatalogFunction)

        self.createLabelAtPosition(1,0,"Path Download: ",50,20)
        self.createLabelAtPosition(1,1,self.supplierInfo['download_path'])
        self.createLabelAtPosition(1,2,"         ")
        self.createButtonAtPosition(1,3,"Browse")
        self.addNormalCommandToButton(1,self.browseDownloadFunction)

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

    def browseCatalogFunction(self):

        self.master.filename  = filedialog.askopenfilename(initialdir=self.supplierInfo['catalogue_path'].split('\\')[:-1], title="Select Catalog", filetypes = (("all files","*.*"),("xlsx files","*.xlsx"),("xls files","*.xls")))
        if self.master.filename == "":
            return
        modifyJson("Kpop","catalogue_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(1,self.master.filename.replace('/','\\'))


    def browseFolderFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['save_path'], title="Select Save Folder")
        if self.master.filename == "":
            return
        modifyJson("Kpop","save_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(7,self.master.filename.replace('/','\\'))


    def browseDownloadFunction(self):

        self.master.filename  = filedialog.askdirectory(initialdir=self.supplierInfo['download_path'], title="Select Download Folder")
        if self.master.filename == "":
            return
        modifyJson("Kpop","download_path",self.master.filename.replace('/','\\'))
        self.changeLabelText(4,self.master.filename.replace('/','\\'))


    def importFileFromWebsite(self):

        driver = webdriver.Chrome(MainApplication.univPath+'\\chromedriver\\chromedriver.exe')  # Optional argument, if not specified will search path.
        driver.get('http://btb.cnlmusic.kr/btb/bbs/login.php');
        time.sleep(0.5)
        driver.find_element_by_id("login_id").send_keys("NICHE")
        driver.find_element_by_id("login_pw").send_keys("Horia@160")
        driver.find_element_by_xpath("//*[@id=\"mb_login\"]/div/div/div/section/form/div[3]/input").click()
        time.sleep(0.5)
        driver.find_element_by_xpath("//*[@id=\"sidebar-menu\"]/div/ul/li[2]/a").click()
        time.sleep(0.5)
        driver.find_element_by_xpath("/html/body/div[1]/div/div[3]/div/div[2]/div/div/div[1]/div[2]/a").click()

        #functioneaza cum trebuie atat timp cat nu este un xlsx sau xls la top la downloads

        while getExtension(newest(self.supplierInfo['download_path'])) != "xls" and  getExtension(newest(self.supplierInfo['download_path'])) != "xlsx":
            time.sleep(1)

        driver.quit()

        file = newest(self.supplierInfo['download_path'])
        print(file)
        print(MainApplication.univPath+"\\temp\\"+getName(file))
        shutil.move(file, MainApplication.univPath+"\\temp\\"+getName(file))


    def solve(self):
        self.initJsonStuff()
        self.importFileFromWebsite()

        """
        error = open(MainApplication.univPath+"\\Resources"+"\\error.txt","w")
        PRICE_ERROR = 696969696969
        BARCODE_ERROR = 797979797979

        file_catalog = self.supplierInfo['catalogue_path']
        file_price = self.supplierInfo['price_path']
        save_path = self.supplierInfo['save_path']
        start_row = self.supplierInfo['start_row']
        price_start_row = self.supplierInfo['price_start_row']
        barcode_column = self.supplierInfo['barcode_column']
        price_column = self.supplierInfo['price_column']
        pricecode_column = self.supplierInfo['pricecode_column']
        rounded_price_column = self.supplierInfo['rounded_price_column']
        save_name = self.supplierInfo['save_name']

        catalogExt = getExtension(file_catalog)
        priceExt = getExtension(file_price)

        piasCatalog = SupplierFile(file_catalog,catalogExt,start_row)
        piasPrices = SupplierFile(file_price,priceExt,price_start_row)

        dictPreturi = piasPrices.getDictionary(pricecode_column,rounded_price_column)

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


        error.close()
        void_workbook.save(save_path+"\\" + save_name)

        self.master.destroy()
        """
