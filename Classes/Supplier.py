
if __name__ == "__main__":
    from MainApplication import *
else:
    from Classes.MainApplication import *

import os
import json

class Supplier(MainApplication):

    jsonFilePath = MainApplication.univPath+"\\Resources\\paths.json"

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

        self.nrButtons = 0
        self.nrEntries = 0
        self.nrLabels = 0

        self.supplierInfo = {}

        with open(Supplier.jsonFilePath) as f:
            data = json.load(f)

        for state in data["suppliers"]:
            if state.get('title') == title:
                self.supplierInfo = state
                break

        if self.supplierInfo['catalogue'] != "null":
            self.cataloguePath = self.supplierInfo['catalogue']

        if self.supplierInfo['price'] != "null":
            self.pricePath = self.supplierInfo['price']

        if self.supplierInfo['save'] != "null":
            self.savePath = self.supplierInfo['save']

        if self.supplierInfo['barcode_column'] != "null":
            self.barcodeColumn = self.supplierInfo['barcode_column']

        if self.supplierInfo['price_column'] != "null":
            self.priceColumn = self.supplierInfo['price_column']

        if self.supplierInfo['start_line'] != "null":
            self.startLine = self.supplierInfo['start_line']






