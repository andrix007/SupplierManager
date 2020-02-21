from Classes.MainApplication import *
from Classes.Supplier import *
import tkinter as tk
import os


supplierTitles = [
"Sony",
"Nuclear Blast",
"Pias",
"Pias_Classical",
"Music On Vynil",
"Sperakers Corner",
"Mystic",
"Sincron",
"Kpop",
"Noutati_Kpop"]


root = tk.Tk()

MainApplication.changeUnivBackColor("black")
MainApplication.changeUnivForColor("white")
MainApplication.changeUnivActiveForColor("green") #for ex. the dropdown menu elements

MainApplication.printPath()

app = MainApplication(root,"SupplierListManagement - NICHE RECORDS S.R.L.©",400,60,r"Icons/nichelogo.ico",MainApplication.univBackColor)
app.createLabelAtPosition(0,0,"Supplier: ",50,20)

app.createDropdownMenuAtPosition(supplierTitles,0,1)
app.createLabelAtPosition(0,2,"           ")
app.createButtonAtPosition(0,3)
app.addNormalCommandToButton(0,app.destroySelf)

root.mainloop()



root = tk.Tk()

name = app.getSelection(0)
supplier = Supplier(root,name,400,60,r"Icons/nichelogo.ico",MainApplication.univBackColor)


root.mainloop()
