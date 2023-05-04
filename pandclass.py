import tkinter as tk            #base import for tkinter
import tkinter.ttk as ttk       #this is for themed widgets 
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import messagebox
import pandas as pd

class ExcelToPandas():
    def __init__(self):
        self.filepath = "error"
        
        
    def pandasfileapprove(self):
        if self.filepath.lower() == "error":
            #print("File path not correctly defined")
            messagebox.askretrycancel(title="File Path Definition Error", message="File path not correctly defined.")
        else: 
            excel_valid = [False, False, False, False]
            file_validator = ["Part Number", "Unit QTY", "QTY", "Description"]
            excel_df = pd.read_excel(self.filepath)
            if len(excel_df.columns) == len(file_validator):
                print(excel_df.head())
                for i in range(len(excel_df.columns)):
                    if excel_df.columns[i] != file_validator[i]:
                        excel_valid[i] = False
                        print("File Not Valid")
                    else: 
                        # print("File confirmed Valid")
                        excel_valid[i] = True
                print(f"Validation: {excel_valid}")
                if excel_valid == [True, True, True, True]:
                    return True
                else:
                    return False
            else:
                messagebox.askretrycancel(title="File Invalid", message="Chosen file is not a valid Inventor BOM Export, please try again")
                return False

            
            
    # def gasket_series(self):
        # print("Starting Gasket Count")
        # excel_df = pd.read_excel(self.filepath)
        # gasket_locations = []
        # excel_df.dropna()
        # df2 = excel_df[excel_df['Part Number'].str.contains("GASKET")]
        # print(df2)
            
            
        # # gasket_dict = {}

        # # print(len(excel_df[excel_df['Part Number'] == "GASKET"]))