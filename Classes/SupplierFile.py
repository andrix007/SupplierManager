if __name__ == "__main__":
    from MainApplication import *
else:
    from Classes.MainApplication import *
if __name__ == "__main__":
    from functions import *
else:
    from Classes.functions import *

class SupplierFile(MainApplication):
    def __init__(self,path=None,extension=None,start_row=None,separator=None):

        if extension == 'csv':
            with open(path, newline = '') as f:
                reader = csv.reader(f, delimiter = separator)
                data = list(reader)
            data = data[1:]
            self.data = data
            return

        self.formula = ""
        self.data = []

        i = 0
        if extension == None:
            self.extension = 'xlsx'
        else:
            self.extension = extension
        if start_row == None:
            self.start_row = 1
        else:
            self.start_row = start_row
        if path == None:
            self.path = "no path"
        else:
            self.path = path

        if extension == None:
            return

        self.wb = xlrd.open_workbook(path)
        self.ws = self.wb.sheet_by_index(0)

        for row in range(self.ws.nrows):
            i = i + 1
            if i < start_row:
                continue

            row_data = []

            for col in range(self.ws.ncols):
                val = self.ws.cell(row,col).value
                row_data.append(val)

            self.data.append(row_data)

    def getBarcodeDictionary(self,barcodeColumn):

        barcodeColumn = barcodeColumn - 1
        dictBar = {}

        for row in self.data:

            value = row[barcodeColumn]
            value = normalizeBarcode(value)
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

    def addOtherSupplierFileInRegardToDictionary(self,otherFile,dictionary,column):

        for row in otherFile.data:
            value = row[column-1]
            if value not in dictionary:
                self.data.append(row)
                dictionary.update({value:"1"})

    def addOtherSupplierFile(self,otherFile,*ignoreColumns):

        for row in otherFile.data:

            nrow = []
            for j in range(len(row)):
                if j in ignoreColumns:
                    continue
                else:
                    nrow.append(row[j])

            self.data.append(nrow)

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

    def printColumn(self, column):

        column = column - 1
        for row in self.data:
            print(row[column-1])


    def addContentToWorkbook(self,workbook_path,start_row):

        wb = load_workbook(workbook_path)
        ws = wb.active

        i = start_row-1

        for row in self.data:

            i = i + 1
            for j in range(1,len(row)+1):
                ws.cell(row = i,column = j).value = row[j-1]

        wb.save(workbook_path)

    def isColumnEmpty(self, column):

        for row in self.data:

            value = row[column-1]
            if str(value).isalpha() or str(value).isdigit():
                return False

        return True

