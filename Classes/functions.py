if __name__ == "__main__":
    from MainApplication import *
else:
    from Classes.MainApplication import *

#File Management Methods Methods <------------------------------------->



#xlrd (for useless very old fcking piece of garbage xls files) OK BOOMER <------------------------------------->

#xlrd (for useless very old fcking piece of garbage xls files) OK BOOMER<------------------------------------->!

#openpyxl (for the modern-ish xlsx files)<------------------------------------->

#openpyxl (for the modern-ish xlsx files)<------------------------------------->!


#other stuff<------------------------------------->

def correctName(name):
    nr = ""
    namae = ""
    for ch in name:
        if ch >= '0' and ch <= '9':
            nr = nr + str(ch)
        else:
            namae = namae + str(ch)

    nr = nr.zfill(2)
    print(namae+nr)
    return namae+nr

def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

def getFileXFromPath(path,x):
    files = []

    orig = os.getcwd()

    os.chdir(path)

    for r,d,f in os.walk("."):
        for file in f:
            if '.xls' in file:
                files.append(os.path.join(r, file))
            elif '.xlsx' in file:
                files.append(os.path.join(r, file))

    cnt = 0

    for f in files:
        cnt = cnt + 1
        if cnt == x:
            os.chdir(orig)
            return path+'\\'+f[2:]


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

def getExtension(file):

    extension = os.path.splitext(file)[1][1:]
    return extension

def convertAll(path):

    files = []

    for folder, subs, filees in os.walk(path):
        for filename in filees:
            files.append(os.path.abspath(os.path.join(folder, filename)))

    for f in files:
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


#other stuff<------------------------------------->!

#File Management Methods Methods <------------------------------------->!