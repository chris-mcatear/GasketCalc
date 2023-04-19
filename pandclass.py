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
            print(excel_df.head())
            for i in range(len(excel_df.columns)):
                if excel_df.columns[i] != file_validator[i]:
                    excel_valid[i] = False
                    print("File Not Valid")
                else: 
                    print("File confirmed Valid")
                    excel_valid[i] = True
            print(excel_valid)
            # excel_df_column_titles = excel_df.columns
            # print(f"colums index: {excel_df.index}")
            # print(f"column titles: {excel_df_column_titles}")
            #expected_column_title = ["Part Number", "Unit QTY", "QTY", "Description"]
            # for column in excel_df_column_titles:
            #     if excel_df_column_titles[column] != expected_column_title[column]:
            #         print("Error, file headers don't match")
            #TODO: Verify the excel file chosen by user by matching column headers ?? idk if this is the best way to go about it 
            
    
    #def print