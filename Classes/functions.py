if __name__ == "__main__":
    from MainApplication import *
else:
    from Classes.MainApplication import *

import time
#File Management Methods Methods <------------------------------------->



#xlrd (for useless very old fcking piece of garbage xls files) OK BOOMER <------------------------------------->

#xlrd (for useless very old fcking piece of garbage xls files) OK BOOMER<------------------------------------->!

#openpyxl (for the modern-ish xlsx files)<------------------------------------->

#openpyxl (for the modern-ish xlsx files)<------------------------------------->!


#other stuff<------------------------------------->
def decode_utf8(string):
    return string.decode('utf-8','replace')

def correctName(name):
    nr = ""
    namae = ""
    for ch in name:
        if ch >= '0' and ch <= '9':
            nr = nr + str(ch)
        else:
            namae = namae + str(ch)

    nr = nr.zfill(2)
    return namae+nr



def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False



def getFileXFromPath(path, x):
    cnt = 0

    for root, dirs, files in os.walk(path):
        for name in files:
            full_path = os.path.join(root, name)
            ext = getExtension(full_path)
            if ext != "xls" and ext != "xlsx" and ext != "csv" and ext != "xlsm":
                return -1 #not a file type I want in the folder
            cnt += 1
            if cnt == x:
                return full_path #Successful execution

    return -2 #Not enought files


def getVersion(file):
    s = file[-7:]
    s = str(s)
    string = ""
    for i in range(len(s)):
        if str(s[i]).isdigit():
            string = string + str(s[i])
    return int(string)

def exist(path):
    orig = os.getcwd()

    os.chdir(path)
    lst = os.listdir(path)
    os.chdir(orig)

    if not lst:
        return False
    else:
        return True



def eraseContent(path):

    for root, dirs, files in os.walk(path):
        for f in files:
            os.unlink(os.path.join(root, f))
        for d in dirs:
            shutil.rmtree(os.path.join(root, d))



def eraseDirectory(path):
    shutil.rmtree(path)



def convertXlsToXlsx(file):

    if getExtension(file) != "xls":
        return

    f = xlrd.open_workbook(file)

    sheet = f.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols

    new_f = openpyxlWorkbook()
    sheet1 = new_f.get_active_sheet()

    for row in range(0, nrows):
        for col in range(0, ncols):
            sheet1.cell(row=row+1, column=col+1).value = sheet.cell_value(row, col)

    new_f.save(file+'x')

    os.remove(file)

def qualityConvertXlsToXlsx(file):

    if getExtension(file) != "xls":
        return

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file)

    wb.SaveAs(file+"x", FileFormat = 51)
    wb.Close()

    os.remove(file)

    excel.Application.Quit()

def uncorruptFile(file):

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file)

    wb.SaveAs(file, FileFormat = 51)
    wb.Close()

    os.remove(file)

    excel.Application.Quit()

def getExtension(file):

    extension = os.path.splitext(file)[1][1:]
    return extension



def convertAll(path):

    files = []

    for folder, subs, filees in os.walk(path):
        for filename in filees:
            files.append(os.path.abspath(os.path.join(folder, filename)))

    for f in files:
        try:
            qualityConvertXlsToXlsx(f)
        except:
            convertXlsToXlsx(f)



def iter_rows(ws):
    for row in ws.iter_rows():
        yield [cell.value for cell in row]



def addFileToOtherFile(file,other_file,coloanaBarcode,randInceput = 2,randTitluri = 1): #aceasta functie presupune ca ambele au aceeasi extensie (.xls/.xlsx) asa ca trebuie apelata mai intai functia de conversie

    orig = os.getcwd()


    other_file_workbook = load_workbook(other_file)
    other_file_sheet = other_file_workbook.active

    file_workbook = load_workbook(file)
    file_sheet = file_workbook.active

    dictBarcode = {}

    for row in list(iter_rows(file_sheet)):
        barcode = row[coloanaBarcode-1]
        if barcode not in dictBarcode:
            dictBarcode.update({barcode:"1"})

    for row in list(iter_rows(other_file_sheet)):
        barcode = row[coloanaBarcode-1]
        if barcode not in dictBarcode:
            dictBarcode.update({barcode:"1"})
            file_sheet.append(row)

    file_workbook.save(filename = file)

    os.remove(other_file)

    os.chdir(orig)



def mergeFiles(pathInitial,pathSalvare,coloanaBarcode,randInceput = 2,randInceputTitluri = 1):

    #delete files after using them

    orig = os.getcwd()

    os.chdir(pathInitial)

    files = []

    #TO DO verifica cu getExtension(file) daca fisierele la care formezi path-ul absolut sunt doar de tipurile xls si xlsx

    for folder, subs, filees in os.walk(pathInitial):
        for filename in filees:
            files.append(os.path.abspath(os.path.join(folder, filename)))

    #print(files) aici in files avem path-urile absolute ale tuturor fisierelor dintr-un directoriu (ar trebuie sa fie toate xls si xlsx, dar trebuie o verificare in plus)

    os.chdir(orig)



def modifyJson(supplierName,element,value):
    supplierInfo = {}

    with open(MainApplication.jsonFilePath) as f:
        data = json.load(f)
    i = 0


    for state in data["suppliers"]:
        if state.get('title') == supplierName:
            data["suppliers"][i].update({element:value})
            break
        i = i + 1

    json_object = json.dumps(data, indent = 4)

    with open(MainApplication.jsonFilePath, "w") as outfile:
        outfile.write(json_object)



def getPriceDict(file,pricecode_column,pricevalue_column,price_start_row):
    dictPreturi = {}
    i = 0

    file_workbook = openpyxlWorkbook(file)
    file_sheet = file_workbook.active

    for row in list(iter_rows(file_sheet)):
        i = i + 1

        if i < price_start_row:
            continue

        pricecode = row[pricecode_column-1]
        pricevalue = row[pricevalue_column-1]

        if pricecode not in dictPreturi:
            dictPreturi.update({pricecode:pricevalue})

    return dictPreturi

def getFormulaForRowX(formula, formula_row, x):
    new_formula = formula.replace("F"+str(formula_row),"F"+str(x)).replace("E"+str(formula_row),"E"+str(x))
    return new_formula


def normalizeBarcode(barcode):

    barcode = str(barcode)

    string = ""
    for i in range(len(barcode)):
        if str(barcode[i]).isdigit():
            string = string + str(barcode[i])

    if '.' not in barcode:
        return string
    try:
        a,b = barcode.split('.')
        return a
    except:
        print("BOI YU IDIOT")
#other stuff<------------------------------------->!


def newest(path):

    if not os.listdir(path):
        return -1

    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files]
    return max(paths, key=os.path.getctime)


def getName(path):
    return os.path.basename(path)
#File Management Methods Methods <------------------------------------->!

def getCorrectPrice(price):
    price = str(price)
    string = ""
    for i in range(len(price)):
        if str(price[i]).isdigit():
            string = string + str(price[i])
    return int(string)

def getCorrectFormat(f):
    f = str(f)
    return f.replace(' ','')


def logText(message, messageColor=MainApplication.univPopupColor):

    popup = tk.Tk()
    popupWindow = MainApplication(popup,"Log",None,None,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    popupWindow.createLabelAtPosition(0, 0, message, 10, 10, None, messageColor)
    popup.mainloop()

def logError(errorMessage, errorColor=MainApplication.univErrorPopupColor):

    popup = tk.Tk()
    popupWindow = MainApplication(popup,"Log",None,None,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    popupWindow.createLabelAtPosition(0, 0, errorMessage, 10, 10, None, errorColor)
    popup.mainloop()

def fileCount(path):
    return len(os.listdir(path))

