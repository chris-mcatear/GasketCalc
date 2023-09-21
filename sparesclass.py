import pandas as pd

class SpareSplitter():
    def __init__(self):
        self.filename = ""
        
    def specials_list(self, master_list):
        excel_df = pd.read_excel(self.filename)
        specials_list = master_list[master_list["Filename"].apply(lambda x: "- GAS" not in x and "- OIL" not in x and "- CW" not in x and "- SEAL" not in x and "- INSULATING" not in x)]
        specials_list['Description'] = specials_list['Description'].fillna(" ")
        # print(f'Specials Gaskets List:\n {specials_list}')
        return specials_list
    
    def master_list(self):
        excel_df = pd.read_excel(self.filename)
        gaskets_master = excel_df[excel_df["Filename"].str.contains("GASKET", na=False, case=False)]
        # print(f'Master Gaskets List:\n {gaskets_master}')
        return gaskets_master