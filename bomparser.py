#BoM Parser by Christopher McAtear
#First commit 08-02-2023 
#Initialising repo for Project. 
#Testing Excel sheet to be stored in folder one higher than project code on local HDD.
#Use CSV module for Python to read information from file, Pandas/Numpy for calculations (possibly)


#Psuedocode
#Run program
#Opens to window which has file select option, option to enter file directory or option to browse PC to select
#Default BoM Layout for input should be standard output that has "Part No, Unit QTY, QTY, Description" in first row. 
#Check to see if file is valid, compare row 1 values to ensure
#Takes standard output BoM from Inventor, scans and counts each part

#Window open use;
#Open window
#Display includes; bar to browse for file, option to drag and drop file, current version, run button

import tkinter as tk            #base import for tkinter
import tkinter.ttk as ttk       #this is for themed widgets 
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import messagebox
from pandclass import ExcelToPandas
    
#Define window
window = tk.Tk()
window.title('BoM Parser - ALPHA 0.0.1')
# window.resizable(False, False)
window.geometry("500x150")

#Defining Style of Window
#style = darkstyle(window)

etop = ExcelToPandas()

def browsefunc():
    filetypes = (("Excel File", "*.xlsx"),)
    filename = fd.askopenfilename(title="Select a file", filetypes=filetypes)
    if filename == "":
        pass
        # if messagebox.askretrycancel(title="Error", message="Please select a file.") == True:
        #     filename = fd.askopenfilename(filetypes=filetypes)
        # else:
        #     pass
    else:
        # showinfo(title="Selected", message = filename)
        etop.filepath = filename
        filepathtext.config(text=filename)
        #showinfo(title="ExcelToPandas File Path", message=etop.filepath)


def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        window.destroy()


def calculatefunc(filename):
    """Usage: calculatefunc(input) / Will take input of file path and pass to pandas to interperetation, pandas to return details to display() function for displaying information."""
    if len(filename) > 0:
        if etop.pandasfileapprove() == True:
            print("Calc func success")
            etop.gasket_series()
        else:
            print("Calc func Failure")   


# Widgets    
filepathframe = ttk.LabelFrame(window, text="Choose a file")
greeting = ttk.Label(text="Please choose a file")
browserbutton = ttk.Button(filepathframe, text="Browse", command=browsefunc)
calculatebutton = ttk.Button(text="Calculate", command= lambda: calculatefunc(etop.filepath))
filepathtext = ttk.Label(filepathframe, text="test")
# retrybutton = ttk.Button(text="Retry", command=browsefunc)

# Content Layout in window
greeting.grid(row=0, column=0, padx=5, pady=5)
filepathframe.grid(row=1, column=0, padx=10, pady=5)
filepathtext.grid(row=0, column=0, padx=5, pady=5)
browserbutton.grid(row=1, column=1, padx=10, pady=5)
calculatebutton.grid(row=1, column=3, padx=10, pady=5)

#Window Remains on Screen
#window.protocol("WM_DELETE_WINDOW", on_closing)

window.mainloop()
