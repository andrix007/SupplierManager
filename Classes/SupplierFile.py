if __name__ == "__main__":
    from MainApplication import *
else:
    from Classes.MainApplication import *
if __name__ == "__main__":
    from functions import *
else:
    from Classes.functions import *

class SupplierFile(MainApplication):
    def __init__(self,path,extension,start_row):

        self.formula = ""
        self.data = []

        i = 0

        self.extension = extension
        self.start_row = start_row
        self.path = path

        self.wb = xlrd.open_workbook(path)
        self.ws = self.wb.sheet_by_index(0)

        for row in range(self.ws.nrows):
            i = i + 1
            if i < start_row:
                continue

            row_data = []

            for col in range(self.ws.ncols):
                row_data.append(self.ws.cell(row,col).value)

            self.data.append(row_data)

    def getBarcodeDictionary(self,barcodeColumn):

        barcodeColumn = barcodeColumn - 1
        dictBar = {}

        for row in self.data:

            value = row[barcodeColumn]

            if value not in dictBar:
                dictBar.update({value:"1"})

        return dictBar

    def getDictionary(self,keyColumn,valueColumn):

        keyColumn = keyColumn - 1
        valueColumn = valueColumn - 1
        dictValues = {}

        for row in self.data:

            key = row[keyColumn]
            value = row[valueColumn]

            if value not in dictValues:
                dictValues.update({key:value})

        return dictValues

    def addOtherSupplierFile(self,otherFile):

        for row in otherFile.data:
            self.data.append(row)

    def getFormula(self,Frow,Fcolumn):

        if self.formula != "":
            return self.formula

        if self.extension == 'xls':
            qualityConvertXlsToXlsx(self.path)
            self.extension = 'xlsx'

        wb = load_workbook(self.path)
        ws = wb.active

        self.formula = ws.cell(row = Frow,column = Fcolumn).value
        return self.formula

    def printData(self):

        for row in self.data:
            print(row)



