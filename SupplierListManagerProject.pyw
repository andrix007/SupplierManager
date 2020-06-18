from Classes.MainApplication import *
from Classes.sony import *
from Classes.nuke import *
from Classes.pias import *
from Classes.Pias_Classical import *
from Classes.MOV import *
from Classes.SCR import *
from Classes.mystic import *
from Classes.sincron import *
from Classes.kpop import *
import tkinter as tk
import os

supplierTitles = [
"Sony",
"Nuclear Blast",
"Pias",
"Pias_Classical",
"Music On Vinyl",
"Speakers Corner",
"Mystic",
"Sincron",
"Kpop",
]



def solve():

    root = tk.Tk()

    name = app.getSelection(0)

    if name == "Sony":
        supplier = Sony(root,name,1035,180,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Nuclear Blast":
        supplier = Nuke(root,name,1100,240,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Pias":
        supplier = Pias(root,name,1035,240,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Pias_Classical":
        supplier = Pias_Classical(root,name,1035,240,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Music On Vinyl":
        supplier = MOV(root,name,1135,180,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Speakers Corner":
        supplier = SCR(root,name,1035,240,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Mystic":
        supplier = Mystic(root,name,1135,240,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Sincron":
        supplier = Sincron(root,name,1055,180,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    elif name == "Kpop":
        supplier = Kpop(root,name,1055,300,r"Icons/nichelogo.ico",MainApplication.univBackColor)
    else:
        print("BOI")

    root.mainloop()

MainApplication.initPath(os.getcwd())
MainApplication.initJsonPath(os.getcwd())
MainApplication.changeUnivBackColor("black")
MainApplication.changeUnivForColor("white")
MainApplication.changeUnivActiveForColor("green") #for ex. the dropdown menu elements
MainApplication.changeUnivPopupColor("#36D54A")
MainApplication.changeUnivErrorPopupColor("#FE0000")

root = tk.Tk()

app = MainApplication(root,"SupplierListManagement - NICHE RECORDS S.R.L.©",420,60,r"Icons/nichelogo.ico",MainApplication.univBackColor)
app.createLabelAtPosition(0,0,"Supplier: ",50,20)

app.createDropdownMenuAtPosition(supplierTitles,0,1)
app.createLabelAtPosition(0,2,"                   ")
app.createButtonAtPosition(0,3)
app.addNormalCommandToButton(0,solve)

root.mainloop()


