import tkinter as tk            #base import for tkinter
import tkinter.ttk as ttk       #this is for themed widgets 
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import messagebox
import pandas as pd
from openpyxl import *
from xlsxwriter import *
from flange_info import onefifty_flange_dict, threehundred_flange_dict

END_RESULT_DF = pd.DataFrame()

OIL_ONE_MATERIAL_CHOICE = "ERROR"
OIL_TWO_MATERIAL_CHOICE = "ERROR"

GAS_ONE_MATERIAL_CHOICE = "ERROR"
GAS_TWO_MATERIAL_CHOICE = "ERROR"

CW_MATERIAL_CHOICE = "ERROR"

ISOLATING_MATERIAL_CHOICE = "ERROR"
SEAL_MATERIAL_CHOICE = "ERROR"
CONDENSATE_MATERIAL_CHOICE = "ERROR"

BOLT_OIL_ONE_MATERIAL_CHOICE = "ERROR"
BOLT_OIL_TWO_MATERIAL_CHOICE = "ERROR"

BOLT_GAS_ONE_MATERIAL_CHOICE = "ERROR"
BOLT_GAS_TWO_MATERIAL_CHOICE = "ERROR"

BOLT_CW_MATERIAL_CHOICE = "ERROR"

BOLT_ISOLATING_MATERIAL_CHOICE = "ERROR"
BOLT_SEAL_MATERIAL_CHOICE = "ERROR"
BOLT_CONDENSATE_MATERIAL_CHOICE = "ERROR"

BOLT_OIL_ONE_WASHER_CHBOX = "ERROR"
BOLT_OIL_TWO_WASHER_CHBOX = "ERROR"
BOLT_GAS_ONE_WASHER_CHBOX = "ERROR"
BOLT_GAS_TWO_WASHER_CHBOX = "ERROR"
BOLT_CW_WASHER_CHBOX = "ERROR"
BOLT_ISOLATING_WASHER_CHBOX = "ERROR"
BOLT_CONDENSATE_WASHER_CHBOX = "ERROR"
BOLT_SEAL_WASHER_CHBOX = "ERROR"


class ExcelToPandas():
    def __init__(self):
        self.filepath = ""
        
        
    def pandasfileapprove(self):
        if self.filepath.lower() == "error":
            #print("File path not correctly defined")
            messagebox.askretrycancel(title="File Path Definition Error", message="File path not correctly defined.")
        else: 
            excel_valid = [False, False, False, False]
            file_validator = ["Item", "Filename", "QTY", "Description"]
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


    def oil_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        oil_gaskets_master_1 = excel_df[excel_df["Filename"].str.contains("OIL 1", na=False)]
        oil_gaskets_master_2 = excel_df[excel_df["Filename"].str.contains("OIL 2", na=False)]
        
        grouped_oil_gaskets_1 = oil_gaskets_master_1.groupby(oil_gaskets_master_1["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        grouped_oil_gaskets_1['Description'] = grouped_oil_gaskets_1['Description'].apply(lambda x: x.split('&&&')[0])
        
        grouped_oil_gaskets_2 = oil_gaskets_master_2.groupby(oil_gaskets_master_2["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        grouped_oil_gaskets_2['Description'] = grouped_oil_gaskets_2['Description'].apply(lambda x: x.split('&&&')[0])
        
        return grouped_oil_gaskets_1, grouped_oil_gaskets_2
        
        
    def gas_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        gas_gaskets_master_1 = excel_df[excel_df["Filename"].str.contains("GAS 1", na=False)]
        gas_gaskets_master_2 = excel_df[excel_df["Filename"].str.contains("GAS 2", na=False)]

        # MERGE MASTER GASKET LIST BY Filename 
        grouped_gas_gaskets_1 = gas_gaskets_master_1.groupby(gas_gaskets_master_1["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        grouped_gas_gaskets_1['Description'] = grouped_gas_gaskets_1['Description'].apply(lambda x: x.split('&&&')[0])
        
        grouped_gas_gaskets_2 = gas_gaskets_master_2.groupby(gas_gaskets_master_2["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        grouped_gas_gaskets_2['Description'] = grouped_gas_gaskets_2['Description'].apply(lambda x: x.split('&&&')[0])
        
        return grouped_gas_gaskets_1, grouped_gas_gaskets_2
    
    
    # Cooling water is C.W. ?? 
    def water_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        cw_gaskets = excel_df[excel_df["Filename"].str.contains("- CW", na=False)]
        grouped_cw_gaskets = cw_gaskets.groupby(cw_gaskets["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        grouped_cw_gaskets['Description'] = grouped_cw_gaskets['Description'].apply(lambda x: x.split('&&&')[0])
        return grouped_cw_gaskets
    
    
    def seal_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        seal_gaskets = excel_df[excel_df["Filename"].str.contains("SEAL", na=False)]
        grouped_seal_gaskets = seal_gaskets.groupby(seal_gaskets["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        return grouped_seal_gaskets
    
    
    def isolating_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        isolating_gaskets = excel_df[excel_df["Filename"].str.contains("INSULATE", na=False)]
        grouped_isolating_gaskets = isolating_gaskets.groupby(isolating_gaskets["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        return grouped_isolating_gaskets
    
    
    def condensate_gaskets(self):
        excel_df = pd.read_excel(self.filepath)
        condensate_gaskets = excel_df[excel_df["Filename"].str.contains("ISOLATING", na=False)]
        grouped_condensate_gaskets = condensate_gaskets.groupby(condensate_gaskets["Filename"]).agg({'QTY': 'sum', 'Description': '&&&'.join})
        return grouped_condensate_gaskets
    
    
    def df_to_excel(self, merged_export):
        filetypes = (("Excel File", "*.xlsx"),)
        filename = fd.asksaveasfilename(initialdir = "/",title = "Select file",filetypes = (("Excel File", "*.xlsx"),("all files","*.*")))
        print(type(merged_export))
        merged_export.to_excel(f'{filename}.xlsx')
        # os.system(f"start EXCEL.EXE {filename}.xlsx")
        messagebox.showinfo(title="Export Success", message="Export was Successul")
        
        
    def material_types(self):
        window_2 = tk.Tk()
        window_2.title("Material Selection")
        window_2.minsize(height=250, width=500)
        
        def button_press():
            # print(oil_one_dropdown.get())
            global OIL_ONE_MATERIAL_CHOICE
            OIL_ONE_MATERIAL_CHOICE = oil_one_dropdown.get()
            global OIL_TWO_MATERIAL_CHOICE
            OIL_TWO_MATERIAL_CHOICE = oil_two_dropdown.get()
            global GAS_ONE_MATERIAL_CHOICE
            GAS_ONE_MATERIAL_CHOICE = gas_one_dropdown.get()
            global GAS_TWO_MATERIAL_CHOICE
            GAS_TWO_MATERIAL_CHOICE = gas_two_dropdown.get()
            global CW_MATERIAL_CHOICE
            CW_MATERIAL_CHOICE = cw_dropdown.get()
            global ISOLATING_MATERIAL_CHOICE
            ISOLATING_MATERIAL_CHOICE = isolating_dropdown.get()
            global SEAL_MATERIAL_CHOICE
            SEAL_MATERIAL_CHOICE = seal_dropdown.get()
            global CONDENSATE_MATERIAL_CHOICE
            CONDENSATE_MATERIAL_CHOICE = condensate_dropdown.get()

            window_2.destroy()
        
        #Dropdowns 
        options = ["A", 
                   "B",
                   "C",
                   "D",
                   "E",
                   "F",
                   "G", 
                   "J",
                   "K",
                   "L",
                   "M",
                   "N",
                   "P", 
                   "S",
                   "T",
                   "U",
                   "X",]
        
        #OIL 1 DROPDOWN MENU
        oil_one_option_var = StringVar()
        oil_one_dropdown = ttk.Combobox(window_2, textvariable=oil_one_option_var)
        oil_one_dropdown.set("A")
        oil_one_dropdown['values'] = options
        oil_one_dropdown['state'] = "readonly"
        oil_one_dropdown.grid(column=2, row=1)
        
        #OIL 2 DROPDOWN MENU
        oil_two_option_var = StringVar()
        oil_two_dropdown = ttk.Combobox(window_2, textvariable=oil_two_option_var)
        oil_two_dropdown.set("A")
        oil_two_dropdown['values'] = options
        oil_two_dropdown['state'] = "readonly"
        oil_two_dropdown.grid(column=2, row=2)
        
        #GAS 1 DROPDOWN MENU
        gas_one_option_var = StringVar()
        gas_one_dropdown = ttk.Combobox(window_2, textvariable=gas_one_option_var)
        gas_one_dropdown.set("A")
        gas_one_dropdown['values'] = options
        gas_one_dropdown['state'] = "readonly"
        gas_one_dropdown.grid(column=2, row=3)
        
        #GAS 2 DROPDOWN MENU
        gas_two_option_var = StringVar()
        gas_two_dropdown = ttk.Combobox(window_2, textvariable=gas_two_option_var)
        gas_two_dropdown.set("A")
        gas_two_dropdown['values'] = options
        gas_two_dropdown['state'] = "readonly"
        gas_two_dropdown.grid(column=2, row=4)
        
        #CW DROPDOWN MENU
        cw_option_var = StringVar()
        cw_dropdown = ttk.Combobox(window_2, textvariable=cw_option_var)
        cw_dropdown.set("A")
        cw_dropdown['values'] = options
        cw_dropdown['state'] = "readonly"
        cw_dropdown.grid(column=2, row=5)
        
        #Isolating DROPDOWN MENU
        isolating_option_var = StringVar()
        isolating_dropdown = ttk.Combobox(window_2, textvariable=isolating_option_var)
        isolating_dropdown.set("A")
        isolating_dropdown['values'] = options
        isolating_dropdown['state'] = "readonly"
        isolating_dropdown.grid(column=2, row=6)
        
        #Seal DROPDOWN MENU
        seal_option_var = StringVar()
        seal_dropdown = ttk.Combobox(window_2, textvariable=seal_option_var)
        seal_dropdown.set("A")
        seal_dropdown['values'] = options
        seal_dropdown['state'] = "readonly"
        seal_dropdown.grid(column=2, row=8)
        
        #Condensate DROPDOWN MENU
        condensate_option_var = StringVar()
        condensate_dropdown = ttk.Combobox(window_2, textvariable=condensate_option_var)
        condensate_dropdown.set("A")
        condensate_dropdown['values'] = options
        condensate_dropdown['state'] = "readonly"
        condensate_dropdown.grid(column=2, row=7)
        
        #Labels
        prompt_text = ttk.Label(window_2, text="Please select material types: ")
        okay_button = ttk.Button(window_2, text="Okay", command=button_press)
        oil_one_text = ttk.Label(window_2, text="Oil 1 Type: ")
        oil_one_info = ttk.Label(window_2, textvariable=OIL_ONE_MATERIAL_CHOICE)
        oil_two_text = ttk.Label(window_2, text="Oil 2 Type: ")
        cw_one_text = ttk.Label(window_2, text="Cooling Water 1 Type: ")
        cw_two_text = ttk.Label(window_2, text="Cooling Water 2 Type: ")
        gas_one_text = ttk.Label(window_2, text="Gas 1 Type: ")
        gas_two_text = ttk.Label(window_2, text="Gas 2 Type: ")
        isolating_text = ttk.Label(window_2, text="Isolating Type: ")
        condensate_text = ttk.Label(window_2, text="Condensate Type: ")
        seal_text = ttk.Label(window_2, text="Seal Type: ")
                
        sw_text = ttk.Label(window_2, text="A = SS inner CS outer (Non-Asbestos Filler) \nB = SS inner SS outer (Non-Asbestos Filler) \nC = SS inner CS outer Low Stress (Graphite Filler) \nD = SS inner CS outer (Graphite Filler)\nE = SS inner SS outer (Graphite Filler)\nF = 304 SS inner SS outer (Graphite Filler)\nG = SS inner SS outer (Teflon Filler)\nJ = SS inner, SS outer (RPTFE Filler)\nK = SS inner, CS outer (RPTFE Filler)\nL = Kamprofile 316L metal core & integral center ring/graphite covering layer\nM = Super Duplex SS inner, Super Duplex SS outer (Graphite Filler)\nN = Duplex SS inner, Duplex SS outer (Graphite Filler)\nP = Alloy 625 inner, Alloy 625 outer (Graphite Filler)\nS = Lamons Inhibitor Gasket (API6FB)\nT = UNS N08825 - 150 BHN\nX = Bolt Grade and Coating to Contract Specific Instruction SCW-FCS", justify="left")
        b_text = ttk.Label(window_2, text="        ")
        cnaf_text=ttk.Label(window_2, text="A = CAF Substitute i.e. Novous Supra \nB = Graphite Metal Laminate \nC = \nD = Grade X Graphited KLINGERSIL C-4500 \nE = Grade X Non-Graphited KLINGERSIL C-4430")
        
        prompt_text.grid(column=0, columnspan=3, row=0)
        okay_button.grid(column=5, row=100)
        oil_one_text.grid(column=0, row=1)
        # oil_one_info.grid(column=5, row=1)
        oil_two_text.grid(column=0, row=2)
        gas_one_text.grid(column=0, row=3)
        gas_two_text.grid(column=0, row=4)
        cw_one_text.grid(column=0, row=5)
        #cw_two_text.grid(column=0, row=6)
        isolating_text.grid(column=0, row=6)
        condensate_text.grid(column=0, row=7)
        seal_text.grid(column=0, row=8)
        
        sw_text.grid(column=4, row=1, rowspan=8, padx=5)
        # b_text.grid(column=3, row=1)
        cnaf_text.grid(column=5, row=1, rowspan=4, padx=5)


    def ax_number_column(self, merged_export):
    
        ax_numbers_list = []
                
        for index, row in merged_export.iterrows():
            description = row["Description"]
            part_numb = index
            temp_ax = "HI"
            
            # SPIRAL OR OTHER GASKET TYPE 
            if "SPIRAL" in description:
                temp_ax += "42"
            elif "C.N.A.F." in description:
                temp_ax += "40"
            else:
                temp_ax +=" else detected"
            
            # RATING     
            if "#150" in description:
                temp_ax += "150"
            elif "#300" in description:
                temp_ax += "300"
            elif "#600" in description:
                temp_ax += "600"
            else:
                temp_ax += "error"
            
            # DIN PIPE SIZE (INCH SIZE x 25 ROUNDED TO NEAREST)
            if "1 1/2'" in description:
                temp_ax += "040"
            elif "1/2'" in description:
                temp_ax += "015"
            elif "2'" in description:
                temp_ax += "050"
            elif "1'" in description:
                temp_ax += "025"
            elif "3/4'" in description:
                temp_ax += "020"
            elif "3'" in description:
                temp_ax += "075"
            elif "4'" in description:
                temp_ax += "100"
            elif "5'" in description:
                temp_ax += "125"
            elif "6'" in description:
                temp_ax += "150"
            elif "7'" in description:
                temp_ax += "175"
            elif "8'" in description:
                temp_ax += "200"
            elif "9'" in description:
                temp_ax += "225"
            elif "10'" in description:
                temp_ax += "250"
            elif "11'" in description:
                temp_ax += "275"
            elif "12'" in description:
                temp_ax += "300"
            elif "20'" in description:
                temp_ax += "500"
            elif "24'" in description:
                temp_ax += "600"
                
            #material choices
            if "OIL 1" in part_numb:
                temp_ax += OIL_ONE_MATERIAL_CHOICE
            elif "OIL 2" in part_numb:
                temp_ax += OIL_TWO_MATERIAL_CHOICE
            elif "GAS 1" in part_numb:
                temp_ax += GAS_ONE_MATERIAL_CHOICE
            elif "GAS 2" in part_numb:
                temp_ax += GAS_TWO_MATERIAL_CHOICE
            elif "- CW" in part_numb:
                temp_ax += CW_MATERIAL_CHOICE
            elif "- SEAL" in part_numb:
                temp_ax += SEAL_MATERIAL_CHOICE
            elif "- CONDENSATE" in part_numb:
                temp_ax += CONDENSATE_MATERIAL_CHOICE
            
            ax_numbers_list.append(temp_ax)
            
        merged_export["AX Numbers"] = ax_numbers_list
        
        
    def bolt_material_types(self):
        window_2 = tk.Tk()
        window_2.title("Bolt Material Selection")
        window_2.minsize(height=250, width=500)
        
        def button_press():
            print("button pressed")
            # print(oil_one_dropdown.get())
            global BOLT_OIL_ONE_MATERIAL_CHOICE
            BOLT_OIL_ONE_MATERIAL_CHOICE = bolt_oil_one_dropdown.get()
            global BOLT_OIL_TWO_MATERIAL_CHOICE
            BOLT_OIL_TWO_MATERIAL_CHOICE = bolt_oil_two_dropdown.get()
            global BOLT_GAS_ONE_MATERIAL_CHOICE
            BOLT_GAS_ONE_MATERIAL_CHOICE = bolt_gas_one_dropdown.get()
            global BOLT_GAS_TWO_MATERIAL_CHOICE
            BOLT_GAS_TWO_MATERIAL_CHOICE = bolt_gas_two_dropdown.get()
            global BOLT_CW_MATERIAL_CHOICE
            BOLT_CW_MATERIAL_CHOICE = bolt_cw_dropdown.get()
            global BOLT_ISOLATING_MATERIAL_CHOICE
            BOLT_ISOLATING_MATERIAL_CHOICE = bolt_isolating_dropdown.get()
            global BOLT_SEAL_MATERIAL_CHOICE
            BOLT_SEAL_MATERIAL_CHOICE = bolt_seal_dropdown.get()
            global BOLT_CONDENSATE_MATERIAL_CHOICE
            BOLT_CONDENSATE_MATERIAL_CHOICE = bolt_condensate_dropdown.get()
            
            global BOLT_OIL_ONE_WASHER_CHBOX
            BOLT_OIL_ONE_WASHER_CHBOX = bolt_oil_one_chbox_var.get()
            global BOLT_OIL_TWO_WASHER_CHBOX
            BOLT_OIL_TWO_WASHER_CHBOX = bolt_oil_two_chbox_var.get()
            global BOLT_GAS_ONE_WASHER_CHBOX
            BOLT_GAS_ONE_WASHER_CHBOX = bolt_gas_one_chbox_var.get()
            global BOLT_GAS_TWO_WASHER_CHBOX
            BOLT_GAS_TWO_WASHER_CHBOX = bolt_gas_two_chbox_var.get()
            global BOLT_CW_WASHER_CHBOX
            BOLT_CW_WASHER_CHBOX = bolt_cw_chbox_var.get()
            global BOLT_ISOLATING_WASHER_CHBOX
            BOLT_ISOLATING_WASHER_CHBOX = bolt_isolating_chbox_var.get()
            global BOLT_CONDENSATE_WASHER_CHBOX
            BOLT_CONDENSATE_WASHER_CHBOX = bolt_condensate_chbox_var.get()
            global BOLT_SEAL_WASHER_CHBOX
            BOLT_SEAL_WASHER_CHBOX = bolt_seal_chbox_var.get()
            
            window_2.destroy()
        
        #Dropdowns 
        options = ["A", 
                   "B",
                   "C",
                   "D",
                   "E",
                   "F",
                   "G", 
                   "J",
                   "K",
                   "L",
                   "M",
                   "N",
                   "P", 
                   "S",
                   "T",
                   "U",
                   "V",
                   "W",
                   "X",]
        
        #OIL 1 DROPDOWN MENU
        bolt_oil_one_option_var = StringVar()
        bolt_oil_one_dropdown = ttk.Combobox(window_2, textvariable=bolt_oil_one_option_var)
        bolt_oil_one_dropdown.set("A")
        bolt_oil_one_dropdown['values'] = options
        bolt_oil_one_dropdown['state'] = "readonly"
        bolt_oil_one_dropdown.grid(column=2, row=1)
        bolt_oil_one_chbox_var = IntVar()
        bolt_oil_one_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_oil_one_chbox_var, onvalue=1, offvalue=0)
        bolt_oil_one_chbox.grid(column=3, row=1)
        
        #OIL 2 DROPDOWN MENU
        bolt_oil_two_option_var = StringVar()
        bolt_oil_two_dropdown = ttk.Combobox(window_2, textvariable=bolt_oil_two_option_var)
        bolt_oil_two_dropdown.set("A")
        bolt_oil_two_dropdown['values'] = options
        bolt_oil_two_dropdown['state'] = "readonly"
        bolt_oil_two_dropdown.grid(column=2, row=2)
        bolt_oil_two_chbox_var = IntVar()
        bolt_oil_two_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_oil_two_chbox_var, onvalue=1, offvalue=0)
        bolt_oil_two_chbox.grid(column=3, row=2)
        
        #GAS 1 DROPDOWN MENU
        bolt_gas_one_option_var = StringVar()
        bolt_gas_one_dropdown = ttk.Combobox(window_2, textvariable=bolt_gas_one_option_var)
        bolt_gas_one_dropdown.set("A")
        bolt_gas_one_dropdown['values'] = options
        bolt_gas_one_dropdown['state'] = "readonly"
        bolt_gas_one_dropdown.grid(column=2, row=3)
        bolt_gas_one_chbox_var = IntVar()
        bolt_gas_one_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_gas_one_chbox_var, onvalue=1, offvalue=0)
        bolt_gas_one_chbox.grid(column=3, row=3)
        
        #GAS 2 DROPDOWN MENU
        bolt_gas_two_option_var = StringVar()
        bolt_gas_two_dropdown = ttk.Combobox(window_2, textvariable=bolt_gas_two_option_var)
        bolt_gas_two_dropdown.set("A")
        bolt_gas_two_dropdown['values'] = options
        bolt_gas_two_dropdown['state'] = "readonly"
        bolt_gas_two_dropdown.grid(column=2, row=4)
        bolt_gas_two_chbox_var = IntVar()
        bolt_gas_two_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_gas_two_chbox_var, onvalue=1, offvalue=0)
        bolt_gas_two_chbox.grid(column=3, row=4)        
        
        #CW DROPDOWN MENU
        bolt_cw_option_var = StringVar()
        bolt_cw_dropdown = ttk.Combobox(window_2, textvariable=bolt_cw_option_var)
        bolt_cw_dropdown.set("A")
        bolt_cw_dropdown['values'] = options
        bolt_cw_dropdown['state'] = "readonly"
        bolt_cw_dropdown.grid(column=2, row=5)
        bolt_cw_chbox_var = IntVar()
        bolt_cw_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_cw_chbox_var, onvalue=1, offvalue=0)
        bolt_cw_chbox.grid(column=3, row=5)
        
        #Isolating DROPDOWN MENU
        bolt_isolating_option_var = StringVar()
        bolt_isolating_dropdown = ttk.Combobox(window_2, textvariable=bolt_isolating_option_var)
        bolt_isolating_dropdown.set("A")
        bolt_isolating_dropdown['values'] = options
        bolt_isolating_dropdown['state'] = "readonly"
        bolt_isolating_dropdown.grid(column=2, row=6)
        bolt_isolating_chbox_var = IntVar()
        bolt_isolating_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_isolating_chbox_var, onvalue=1, offvalue=0)
        bolt_isolating_chbox.grid(column=3, row=6)
        
        #Seal DROPDOWN MENU
        bolt_seal_option_var = StringVar()
        bolt_seal_dropdown = ttk.Combobox(window_2, textvariable=bolt_seal_option_var)
        bolt_seal_dropdown.set("A")
        bolt_seal_dropdown['values'] = options
        bolt_seal_dropdown['state'] = "readonly"
        bolt_seal_dropdown.grid(column=2, row=8)
        bolt_seal_chbox_var = IntVar()
        bolt_seal_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_seal_chbox_var, onvalue=1, offvalue=0)
        bolt_seal_chbox.grid(column=3, row=8)
        
        #Condensate DROPDOWN MENU
        bolt_condensate_option_var = StringVar()
        bolt_condensate_dropdown = ttk.Combobox(window_2, textvariable=bolt_condensate_option_var)
        bolt_condensate_dropdown.set("A")
        bolt_condensate_dropdown['values'] = options
        bolt_condensate_dropdown['state'] = "readonly"
        bolt_condensate_dropdown.grid(column=2, row=7)
        bolt_condensate_chbox_var = IntVar()
        bolt_condensate_chbox = Checkbutton(window_2, text="Washers?", variable=bolt_condensate_chbox_var, onvalue=1, offvalue=0)
        bolt_condensate_chbox.grid(column=3, row=7)
        
        #Labels
        prompt_text = ttk.Label(window_2, text="Please select material types: ")
        okay_button = ttk.Button(window_2, text="Okay", command=button_press)
        oil_one_text = ttk.Label(window_2, text="Oil 1 Type: ")
        oil_one_info = ttk.Label(window_2, textvariable=OIL_ONE_MATERIAL_CHOICE)
        oil_two_text = ttk.Label(window_2, text="Oil 2 Type: ")
        cw_one_text = ttk.Label(window_2, text="Cooling Water 1 Type: ")
        cw_two_text = ttk.Label(window_2, text="Cooling Water 2 Type: ")
        gas_one_text = ttk.Label(window_2, text="Gas 1 Type: ")
        gas_two_text = ttk.Label(window_2, text="Gas 2 Type: ")
        isolating_text = ttk.Label(window_2, text="Isolating Type: ")
        condensate_text = ttk.Label(window_2, text="Condensate Type: ")
        seal_text = ttk.Label(window_2, text="Seal Type: ")
                
        bolt_text = ttk.Label(window_2, text="A = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Zinc Plated \nB = ASTM A320 L7M Studbolt, ASTM A194 7M(S3) Nuts, Galvanised and Centrifuged to ISO 1461 \nC = ASTM A193-B8M Studbolt, ASTM A194-B8 Nuts, Stainless Steel \nD = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Hot Dip Galvanised to ASTM A153 and A143\nE = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Cadmium Plated ASTM A165 and Bichromatised Yellow ASTM B201\nF = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Cadmium Plated ASTM A165 and Bichromatised Yellow ASTM B201\nG = ASTM A193 B7M Studbolt, ASTM A194 2HM Nuts, Double Spun Galvanised to ASTM A153/A153M\nH = ASTM A193-B8M Studbolt, ASTM A194-B8 Nuts, Stainless Steel \nJ = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Xylan 1070 Coated (Colour Blue)\nK = ASTM A193-B7M Studbolt, ASTM A194-2HM Nuts, Zinc Plated / Xylan 1070 Coated\nL = ASTM A193-B7M Studbolt, ASTM A194-2HM Nuts, Zinc Plated / Xylan 1070 Coated (Colour Blue)\nM = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Carbon Steel - Black (Uncoated)\nN = ASTM A193-B8 Cl2 Studbolt, ASTM A194-B8 Nuts, Stainless Steel - Grade 304\nP = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Spun Galvanised to BS7371 Part 6\nR = ASTM A193-B7 Studbolt, ASTM A194-2H Nuts, Zinc Plated / Xylan 1070 Coated \nS = Inconel 718 A1014-N07718 w/2 Heavy Hex Nuts, B1.1, NACE MR0175 / ISO 15156\nT = Super Duplex SS (S32760) W/2 Heavy Hex Nuts, B1.1, NACE MR0175 / ISO 15156 \nU = Refer to Quality Plan for Specification \nV= ASTM A193-B7M Studbolt, ASTM A194-HM Cermet Flouropolymer Coated W/2 Heavy Hex Nuts, B1.1, NACE MR0175 / ISO 15156 \nW = Refer to Quality Plan for Inspection\nX = Bolt Grade and Coating to Contract Specific Instruction SCW-FCS", justify="left")
        b_text = ttk.Label(window_2, text="        ")
        
        prompt_text.grid(column=0, columnspan=3, row=0)
        okay_button.grid(column=5, row=100)
        oil_one_text.grid(column=0, row=1)
        # oil_one_info.grid(column=5, row=1)
        oil_two_text.grid(column=0, row=2)
        gas_one_text.grid(column=0, row=3)
        gas_two_text.grid(column=0, row=4)
        cw_one_text.grid(column=0, row=5)
        #cw_two_text.grid(column=0, row=6)
        isolating_text.grid(column=0, row=6)
        condensate_text.grid(column=0, row=7)
        seal_text.grid(column=0, row=8)
        
        bolt_text.grid(column=5, row=1, rowspan=8)
        b_text.grid(column=4, row=1)

    def bolt_quantity(self, merged_export):
        bolt_qty_list = []
        bolt_size_list = []
        bolt_length_list = []
        bolt_hpc_num = []
                
        for index, row in merged_export.iterrows():
            description = row["Description"]
            part_numb = index
            flange_qty = row["QTY"]
            
            if "in" in description:
                size = description.split(" ")
                size = size[0]
            elif "NB" in description:
                size = description.split("''NB ")
                size = size[0]      
            
            if "#150" in description:
                bolt_qty_list.append(flange_qty * onefifty_flange_dict[size]['bolt count'])
                bolt_size_list.append(onefifty_flange_dict[size]['bolt size'])
                bolt_length_list.append(onefifty_flange_dict[size]['bolt length'])
                temp_bolt_ax = onefifty_flange_dict[size]['hpc_no']
                # this does work but just need to filter for correct service

            elif "#300" in description:
                bolt_qty_list.append(flange_qty * threehundred_flange_dict[size]['bolt count'])
                bolt_size_list.append(threehundred_flange_dict[size]['bolt size'])
                bolt_length_list.append(threehundred_flange_dict[size]['bolt length'])
                temp_bolt_ax = threehundred_flange_dict[size]['hpc_no']
                
            if "OIL 1" in part_numb:
                temp_bolt_ax += BOLT_OIL_ONE_MATERIAL_CHOICE
                if BOLT_OIL_ONE_WASHER_CHBOX == 0:
                    test_value = bolt_length_list[-1]
                    bolt_length_list[-1] = test_value + 0.5
                    
            elif "OIL 2" in part_numb:
                temp_bolt_ax += BOLT_OIL_TWO_MATERIAL_CHOICE
                if BOLT_OIL_TWO_WASHER_CHBOX == 0:
                    test_value = bolt_length_list[-1]
                    bolt_length_list[-1] = test_value + 0.5
                    
            elif "GAS 1" in part_numb:
                temp_bolt_ax += BOLT_GAS_ONE_MATERIAL_CHOICE
                if BOLT_GAS_ONE_WASHER_CHBOX == 0:
                    test_value = bolt_length_list[-1]
                    bolt_length_list[-1] = test_value + 0.5
                    
            elif "GAS 2" in part_numb:
                temp_bolt_ax += BOLT_GAS_TWO_MATERIAL_CHOICE
                if BOLT_GAS_TWO_WASHER_CHBOX == 0:
                    test_value = bolt_length_list[-1]
                    bolt_length_list[-1] = test_value + 0.5
                    
            elif "- CW" in part_numb:
                temp_bolt_ax += BOLT_CW_MATERIAL_CHOICE
                if BOLT_CW_WASHER_CHBOX == 0:
                    test_value = bolt_length_list[-1]
                    bolt_length_list[-1] = test_value + 0.5
                    
            elif "- SEAL" in part_numb:
                temp_bolt_ax += BOLT_SEAL_MATERIAL_CHOICE
                # if BOLT_SEAL_WASHER_CHBOX == 0:
                #     test_value = bolt_length_list[-1]
                #     bolt_length_list[-1] = test_value + 0.5
                    
            elif "- CONDENSATE" in part_numb:
                temp_bolt_ax += BOLT_CONDENSATE_MATERIAL_CHOICE
                if BOLT_CONDENSATE_WASHER_CHBOX == 0:
                    test_value = bolt_length_list[-1]
                    bolt_length_list[-1] = test_value + 0.5
                    
            elif "- ISOLATING" in part_numb:
                temp_bolt_ax += BOLT_ISOLATING_MATERIAL_CHOICE
                if BOLT_ISOLATING_WASHER_CHBOX == 0:
                    test_value = bolt_length_list[-1]
                    bolt_length_list[-1] = test_value + 0.5
            
            bolt_hpc_num.append(temp_bolt_ax)
            
        merged_export["Bolt Quanities"] = bolt_qty_list
        merged_export["Bolt Size"] = bolt_size_list
        merged_export["Bolt Length"] = bolt_length_list
        merged_export["Bolt HCO No."] = bolt_hpc_num