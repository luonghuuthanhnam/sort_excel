import pandas as pd
import numpy as np
import string
from openpyxl import load_workbook

class Sorting():
    def __init__(self, config_dict) -> None:
        self.config_dict = config_dict
        self.process_sheet_name = config_dict["sheet_name"]
        self.raw_file_path = config_dict["raw_file_path"]
        self.save_file_path = config_dict["save_file_path"]

    def read_sheet_for_process(self):
        excel_file = pd.ExcelFile(self.raw_file_path)
        print(f"Processing Sheet: {self.process_sheet_name}")
        temp_sheet = excel_file.parse(self.process_sheet_name)
        t = temp_sheet[["Unnamed: 6"]].dropna()
        temp = t.loc[t["Unnamed: 6"].str.contains("PortPin")]
        temp["boolean"] = temp["Unnamed: 6"].map( lambda x: x.replace("PortPin","").isdigit())
        temp = temp.reset_index()

        port_group = temp_sheet[["Unnamed: 5"]]
        port_group = port_group.dropna()

        port_gr_temp = port_group.loc[port_group["Unnamed: 5"].str.contains("PortGroup")]
        port_gr_temp = port_gr_temp.reset_index()


        test_df = temp[temp["boolean"]==True]
        test_df = test_df.sort_values("index")

        list_pairs = []
        for i in range(len(test_df) - 1):
            idx, value, _ = test_df.iloc[i]
            next_idx, next_value, _ = test_df.iloc[i+1]
            # portpin_num = value.replace("PortPin","")
            # next_portpin_num = next_value.replace("PortPin","")
            if next_idx-1 not in port_gr_temp["index"].to_list():
                list_pairs.append((idx+1, next_idx))
            else:
                list_pairs.append((idx+1, next_idx-1))
        return list_pairs, temp_sheet

    def append_to_sheet(self, sheet, row_idx, series_row_overwrite):
        real_idx = row_idx + 2
        for idx, char in enumerate(list(string.ascii_uppercase)):
            cell_name = char + str(real_idx)
            sheet[cell_name] = series_row_overwrite[idx]
        sheet["AA" + str(real_idx)] = series_row_overwrite[-2]
        sheet["AB" + str(real_idx)] = series_row_overwrite[-1]
        return sheet

    def Start_Sorting(self, list_pairs, temp_sheet):
        for pair in list_pairs:
            pos_from, pos_to = pair
            sub_df = temp_sheet[pos_from:pos_to]
            sub_df = sub_df.sort_values("Unnamed: 7")

            for r_idx in range(pos_from ,pos_to):
                row = sub_df.iloc[r_idx - pos_from]
                self.sheet = self.append_to_sheet(self.sheet, r_idx, row)

    def open_workbook(self):
        self.workbook = load_workbook(filename=self.raw_file_path)
        sheet = self.workbook[self.process_sheet_name]
        return sheet

    def save_workbook(self):
        self.workbook.save(filename=self.save_file_path)
        self.workbook.close()

    def __call__(self):
        print(f"Opening workbook: {self.raw_file_path} ...")
        self.sheet = self.open_workbook()
        print(f"Sorting Sheet: {self.process_sheet_name} ...")
        list_pairs, temp_sheet = self.read_sheet_for_process()
        self.Start_Sorting(list_pairs, temp_sheet)

        print(f"Saving to: {self.save_file_path} ...")
        self.save_workbook()
        print(f"DONE")
