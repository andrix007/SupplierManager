from Classes.MainApplication import *
from Classes.sony import *
from Classes.nuclear_blast import *
from Classes.pias import *
from Classes.Pias_Classical import *
from Classes.MOV import *
import tkinter as tk
import os

supplierTitles = [
"Sony",
"Nuclear Blast",
"Pias",
"Pias_Classical",
"Music On Vinyl",
"Sperakers Corner",
"Mystic",
"Sincron",
"Kpop",
"Noutati_Kpop"]

##########################################################################################################################################################################################################

MainApplication.initPath(os.getcwd())
MainApplication.initJsonPath(os.getcwd())
MainApplication.changeUnivBackColor("black")
MainApplication.changeUnivForColor("white")
MainApplication.changeUnivActiveForColor("green") #for ex. the dropdown menu elements

root = tk.Tk()

app = MainApplication(root,"SupplierListManagement - NICHE RECORDS S.R.L.©",400,60,r"Icons/nichelogo.ico",MainApplication.univBackColor)
app.createLabelAtPosition(0,0,"Supplier: ",50,20)

app.createDropdownMenuAtPosition(supplierTitles,0,1)
app.createLabelAtPosition(0,2,"                   ")
app.createButtonAtPosition(0,3)
app.addNormalCommandToButton(0,app.destroySelf)

root.mainloop()

##########################################################################################################################################################################################################

root = tk.Tk()

name = app.getSelection(0)

if name == "Sony":
    supplier = Sony(root,name,835,120,r"Icons/nichelogo.ico",MainApplication.univBackColor)
elif name == "Nuclear Blast":
    supplier = Nuclear(root,name,835,180,r"Icons/nichelogo.ico",MainApplication.univBackColor)
elif name == "Pias":
    supplier = Pias(root,name,835,180,r"Icons/nichelogo.ico",MainApplication.univBackColor)
elif name == "Pias_Classical":
    supplier = Pias_Classical(root,name,835,180,r"Icons/nichelogo.ico",MainApplication.univBackColor)
elif name == "Music On Vinyl":
    supplier = MOV(root,name,835,120,r"Icons/nichelogo.ico",MainApplication.univBackColor)
else:
    print("BOI")

root.mainloop()

