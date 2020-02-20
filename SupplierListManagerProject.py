from Classes.MainApplication import *
import tkinter as tk

root = tk.Tk()

MainApplication.changeUnivBackColor("black")
MainApplication.changeUnivForColor("white")

app = MainApplication(root,"SupplierListManagement - NICHE RECORDS S.R.L.©",400,200,r"Icons/nichelogo.ico",MainApplication.univBackColor)
app.createLabelAtPosition(0,0,"Supplier: ",50,20)
options = [
"Sony",
"Nuclear Blast",
"PIAS"]
app.createDropdownMenuAtPosition(options,0,1,"black","red")
app.createLabelAtPosition(0,2,"           ")
app.createButtonAtPosition(0,3)
app.addLambdaCommandToButton(0,app.showSelection,0)

root.mainloop()

