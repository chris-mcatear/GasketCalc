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
import pandas as pd
from pandclass import ExcelToPandas
import pandastable as pt
from pandastable import *
    
#Define window
window = tk.Tk()
window.title('BoM Parser - ALPHA 0.0.1')
# window.resizable(False, False)
# window.geometry("1000x150")
window.config(padx=25, pady=25)
window.minsize(height=500)

#Defining Style of Window
#style = darkstyle(window)

etop = ExcelToPandas()
MATERIAL_CHOSEN = False

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
        # print(filename)
        filepathtext.config(text=filename)
        
        #showinfo(title="ExcelToPandas File Path", message=etop.filepath)
        if len(filename) > 0:
            if etop.pandasfileapprove() == True:
                # print("Calc func success")
                file_approved.config(text="File Valid!", foreground="#11a713")
                calculatebutton.config(state=NORMAL)
                # etop.gasket_series()
            else:
                # print("Calc func Failure")
                file_approved.config(text="File Not Valid! Double check BoM follows order of: Item, Part Number, QTY, Description.", foreground="#f00")
                calculatebutton.config(state=DISABLED)
                # messagebox.askretrycancel(title="File Invalid", message="Chosen file is not a valid Inventor BOM Export, please try again")


def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        window.destroy()


def calculatefunc():
    """Usage: calculatefunc(input) / Will take input of file path and pass to pandas to interperetation, pandas to return details to display() function for displaying information."""
    # etop.gasket_series()
    
    oil_1_gaskets, oil_2_gaskets = etop.oil_gaskets()
    gas_1_gaskets, gas_2_gaskets = etop.gas_gaskets()
    cw_gaskets = etop.water_gaskets()
    seal_gaskets = etop.seal_gaskets()
    isolating_gaskets = etop.isolating_gaskets()
    condensate_gaskets = etop.condensate_gaskets()
    
    merged_gaskets_master = [gas_1_gaskets, gas_2_gaskets, oil_1_gaskets, oil_2_gaskets, cw_gaskets, seal_gaskets, isolating_gaskets, condensate_gaskets]
    final_grouping = pd.concat(merged_gaskets_master)
    # print(gas_1_gaskets)
    # print(gas_2_gaskets)
    # print(cw_gaskets)
    # print(final_grouping)
    return final_grouping
   
    
def export_to_excel():
    # THIS CURRENTLY ISNT EXPORT IT IS SOLEY TO TEST ADDING NEW COLUMN
    merged_export = calculatefunc()
    print(merged_export)
    etop.ax_number_column(merged_export)
    etop.bolt_quantity(merged_export)
    etop.df_to_excel(merged_export)
    
    
def popup_window():
    results_window = Toplevel(window)
    results_window.minsize(height=500, width=1000)
    results_window.title("Results")
    
    final_grouping = calculatefunc()
    # final_grouping.insert(0, 'test', final_grouping)
    
    # frame = tk.Frame(results_window)
    # frame.grid(row=0, column=0)
    
    # tableview = pt.Table(frame, dataframe=final_grouping, showtoolbar=True, showstatusbar=True)
    # tableview.show()
    
    tree = ttk.Treeview(results_window)
    tree["columns"] = tuple(final_grouping.columns)
    
    for column in final_grouping.columns:
        tree.column(column, width=100, anchor="w")
        tree.heading(column, text=column)
        
    for index, row in final_grouping.iterrows():
        tree.insert("", "end", text=index, values=tuple(row))
        
    tree.pack(expand=True, fill="both")


def material_chooser():
    # THIS EXISTS JUST TO CHANGE THE FKN EXCEL BUTTON ASDJDGSGNDSGNDGN
    etop.material_types()
    export_button["state"] = tk.NORMAL
    
    
# Dropdown Options
# options = ["A", 
#            "B",
#            "C",
#            "D",]

# option_var = tk.StringVar()

# Widgets    
filepathframe = ttk.LabelFrame(text="Filepath: ")
greeting = ttk.Label(text="Please choose a file")
browserbutton = ttk.Button(text="Browse", command=browsefunc, width=15)
calculatebutton = ttk.Button(text="Calculate", command=calculatefunc, width=15)
skip_choicebutton = ttk.Button(text="Skip Choice", command=calculatefunc, width=15)
filepathtext = ttk.Label(filepathframe, text="Awaiting file selection.", width=70)
file_approved = ttk.Label(text="Awaiting file selection.")
export_button = ttk.Button(text="Export to Excel", command=export_to_excel, width=15)
export_button["state"] = tk.DISABLED
preview_button = ttk.Button(text="Preview data", command=popup_window, width=15)
#oil_dropdown = ttk.OptionMenu(window, option_var, options[0], *options)
material_window_button = ttk.Button(text="Material Window", command=material_chooser)


# Content Layout in window
filepathframe.grid(row=1, column=0, columnspan=6, padx=25, pady=25)
filepathtext.grid(row=1, column=0, columnspan=3, padx=10, pady=5)
browserbutton.grid(row=0, column=0)
# calculatebutton.grid(row=0, column=1)
file_approved.grid(row=2, column=0, columnspan=6)
# skip_choicebutton.grid(column=2, row=0)
export_button.grid(column=4, row=0)
# preview_button.grid(column=5, row=0)
# oil_dropdown.grid(column=10, row=0)
material_window_button.grid(column=3, row=0)


window.mainloop()
