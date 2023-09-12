#BoM Parser by Christopher McAtear

import tkinter as tk            #base import for tkinter
import tkinter.ttk as ttk       #this is for themed widgets 
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import messagebox
import pandas as pd
from pandclass import *
import pandastable as pt
from pandastable import *


#Define window
window = tk.Tk()
window.title('Gaskets & Bolts Calculator - BETA 1.1')
# window.resizable(False, False)
# window.geometry("1000x150")
window.config(padx=25, pady=25)
window.minsize(height=500)

etop = ExcelToPandas()
MATERIAL_CHOSEN = False


def browsefunc():
    filetypes = (("Excel File", "*.xlsx"),)
    filename = fd.askopenfilename(title="Select a file", filetypes=filetypes)
    if filename == "":
        pass
    else:
        etop.filepath = filename
        filepathtext.config(text=filename)
        
        if len(filename) > 0:
            if etop.pandasfileapprove() == True:
                file_approved.config(text="File Valid!", foreground="#11a713")
                material_window_button.config(state=NORMAL)
            else:
                file_approved.config(text="File Not Valid! Double check BoM follows order of: Item, Part Number, QTY, Description.", foreground="#f00")
                material_window_button.config(state=DISABLED)


def calculatefunc():
    """Usage: calculatefunc(input) / Will take input of file path and pass to pandas to interperetation, pandas to return details to display() function for displaying information."""  
    oil_1_gaskets, oil_2_gaskets = etop.oil_gaskets()
    gas_1_gaskets, gas_2_gaskets = etop.gas_gaskets()
    cw_gaskets = etop.water_gaskets()
    seal_gaskets = etop.seal_gaskets()
    isolating_gaskets = etop.isolating_gaskets()
    condensate_gaskets = etop.condensate_gaskets()
    
    merged_gaskets_master = [gas_1_gaskets, gas_2_gaskets, oil_1_gaskets, oil_2_gaskets, cw_gaskets, seal_gaskets, isolating_gaskets, condensate_gaskets]
    final_grouping = pd.concat(merged_gaskets_master)
    return final_grouping
   
    
def export_to_excel():
    merged_export = calculatefunc()
    etop.ax_number_column(merged_export)
    etop.bolt_quantity(merged_export)
    etop.df_to_excel(merged_export)


def issues_window():
    contact_window = Toplevel(window)
    contact_window.title("Contact")
    
    contact_name = ttk.Label(contact_window, text="Creator: Christopher McAtear")
    contact_email = ttk.Label(contact_window, text="Email: christopher.mcatear@howden.com")
    
    contact_name.grid(column=0, row=0, padx=25, pady=10)
    contact_email.grid(column=0, row=1, padx=25, pady=10)


def material_chooser():
    etop.material_types()
    bolt_material_window_button["state"] = tk.NORMAL
    
    
def bolt_material_chooser():
    etop.bolt_material_types()
    export_button["state"] = tk.NORMAL
    

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
# preview_button = ttk.Button(text="Preview data", command=popup_window, width=15)
#oil_dropdown = ttk.OptionMenu(window, option_var, options[0], *options)
material_window_button = ttk.Button(text="Gasket Materials", command=material_chooser)
material_window_button["state"] = tk.DISABLED
bolt_material_window_button = ttk.Button(text="Bolt Materials", command=bolt_material_chooser)
bolt_material_window_button["state"] = tk.DISABLED
contact_button = ttk.Button(text="Got a Problem?", command=issues_window)


# Content Layout in window
filepathframe.grid(row=1, column=0, columnspan=4, padx=25, pady=25)
filepathtext.grid(row=1, column=0, columnspan=4, padx=10, pady=5)
browserbutton.grid(row=0, column=0)
# calculatebutton.grid(row=0, column=1)
file_approved.grid(row=2, column=0, columnspan=4)
# skip_choicebutton.grid(column=2, row=0)
export_button.grid(column=3, row=0)
# preview_button.grid(column=5, row=0)
# oil_dropdown.grid(column=10, row=0)
material_window_button.grid(column=1, row=0)
bolt_material_window_button.grid(column=2, row=0)
contact_button.grid(column=1, row=5, columnspan=2, pady=10)


window.mainloop()