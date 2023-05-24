import tkinter as tk            #base import for tkinter
import tkinter.ttk as ttk       #this is for themed widgets 
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import messagebox
import pandas as pd

class ExcelToPandas():
    def __init__(self):
        self.filepath = "C:/Users/McAteach/OneDrive - Howden Group Ltd/Coding/BoM Counter/P79 BoM 240523.xlsx"
        
        
    def pandasfileapprove(self):
        if self.filepath.lower() == "error":
            #print("File path not correctly defined")
            messagebox.askretrycancel(title="File Path Definition Error", message="File path not correctly defined.")
        else: 
            excel_valid = [False, False, False, False]
            file_validator = ["Item", "Part Number", "QTY", "Description"]
            excel_df = pd.read_excel(self.filepath)
            if len(excel_df.columns) == len(file_validator):
                # print(excel_df.head())
                for i in range(len(excel_df.columns)):
                    if excel_df.columns[i] != file_validator[i]:
                        excel_valid[i] = False
                        # print("File Not Valid")
                    else: 
                        # print("File confirmed Valid")
                        excel_valid[i] = True
                # print(f"Validation: {excel_valid}")
                if excel_valid == [True, True, True, True]:
                    return True
                else:
                    return False
            else:
                # messagebox.askretrycancel(title="File Invalid", message="Chosen file is not a valid Inventor BOM Export, please try again")
                return False

            
    def gasket_series(self):
        # print("Starting Gasket Count")
        # print(self.filepath)
        excel_df = pd.read_excel(self.filepath)
        # print(type(excel_df))
        # print(excel_df.head())
        gasket_values = excel_df[excel_df["Part Number"].str.contains("GASKET", na=False)]
        # print(gasket_values)
        three_hundred_pound = gasket_values[gasket_values["Part Number"].str.contains("#300", na=False)]
        # print(three_hundred_pound)
        # FROM HERE OUT THREE HUNDRED POUND FLANGE VAR WLL BE SHORTENED TO THP_*NAME* E.G. THP_ONE_INCH
        thp_half_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("1/2'", na=False)]
        thp_three_quart_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("3/4'", na=False)]
        thp_one_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("1'", na=False)]
        thp_one_half_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("1 1/2'", na=False)]
        thp_two_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("2'", na=False)]
        thp_three_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("3'", na=False)]
        thp_six_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("6'", na=False)]
        thp_eight_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("8'", na=False)]
        thp_ten_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("10'", na=False)]
        thp_twelve_inch = three_hundred_pound[three_hundred_pound["Description"].str.contains("12'", na=False)]
        
        one_hundred_fifty_pound = gasket_values[gasket_values["Part Number"].str.contains("#150", na=False)]
        # FROM HERE OUT ONE HUNDRED FIFTY POUND FLANGE VAR WLL BE SHORTENED TO OHFP_*NAME* E.G. OHFP_ONE_INCH
        ohfp_half_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("1/2'", na=False)]
        ohfp_three_quart_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("3/4'", na=False)]
        ohfp_one_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("1'", na=False)]
        ohfp_one_half_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("1 1/2'", na=False)]
        ohfp_two_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("2'", na=False)]
        ohfp_three_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("3'", na=False)]
        ohfp_four_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("4'", na=False)]
        ohfp_six_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("6'", na=False)]
        ohfp_eight_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("8'", na=False)]
        ohfp_ten_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("10'", na=False)]
        ohfp_twelve_inch = one_hundred_fifty_pound[one_hundred_fifty_pound["Description"].str.contains("12'", na=False)]
        
        print(three_hundred_pound)
        print(one_hundred_fifty_pound)
    
    
    def oil_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        oil_gaskets_master = excel_df[excel_df["Part Number"].str.contains("OIL", na=False)]
        print(oil_gaskets_master)
        
        
    def gas_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        gas_gaskets_master_1 = excel_df[excel_df["Part Number"].str.contains("GAS 1", na=False)]
        gas_gaskets_master_2 = excel_df[excel_df["Part Number"].str.contains("GAS 2", na=False)]
        print(gas_gaskets_master_1)
        # print(gas_gaskets_master_2)

        # MERGE MASTER GASKET LIST BY PART NUMBER 
        merged_gas_gaskets_1 = gas_gaskets_master_1.groupby(gas_gaskets_master_1["Part Number"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        merged_gas_gaskets_1['Description'] = merged_gas_gaskets_1['Description'].apply(lambda x: x.split('&&&')[0])
        # merged_gas_gaskets_1 = merged_gas_gaskets_1.sort_values(by="Part Number")
        print(f"\n \n \n \n Merged Gaskets: \n")
        print(merged_gas_gaskets_1)
        return gas_gaskets_master_1, gas_gaskets_master_2
    
    
    # Cooling water is C.W. ?? 
    def water_gaskets(self):
        pass
    
    
    # def df_to_excel(self, gasket_selection):
    #     gasket_selection.to_excel()
    
    # https://stackoverflow.com/questions/46826773/how-can-i-merge-rows-by-same-value-in-a-column-in-pandas-with-aggregation-func
    # this is how to combine dataframes to make them shorter, should be able to add this after master lists and get aggregated lists. 
        