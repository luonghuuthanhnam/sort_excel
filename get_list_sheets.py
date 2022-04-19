import yaml
import pandas as pd

config_dict = {}
with open('config.yaml') as f:
    data = yaml.load(f, Loader=yaml.FullLoader)
    config_dict = dict(data)

excel_file = pd.ExcelFile(config_dict["raw_file_path"])
sheets_name = excel_file.sheet_names
for idx, name in enumerate(sheets_name):
    print(idx, name)