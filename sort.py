from tabnanny import verbose
import pandas as pd
from openpyxl import load_workbook

# start

# Sheet_dict is a dictionary with pair {sheet_name:dataframe}
# all_sheets is a dictionary with all dataframes
sheets_dict = pd.read_excel('data.xlsx', sheet_name=None)

all_sheets = []
for name, sheet in sheets_dict.items():
        # sheet['sheet'] = name
        sheet = sheet.rename(columns=lambda x: x.split('\n')[-1])
        all_sheets.append(sheet)

# Combine all the sheets together to one big sheet
full_table = pd.concat(all_sheets)
full_table.reset_index(inplace=True, drop=True)

full_table.to_excel('test.xlsx')


# Makes a dictionary with pair {sheet:dictionary} and that dictionary with {column_name:column_values} 
sheets_col_dict = {}
for name, sheet in sheets_dict.items():
        sheets_col_dict[name] = []
        for columns in sheet.columns:
                col_dict = {}
                col_dict[str(columns)] = sheet[columns]  
                sheets_col_dict[name] += [col_dict]

print(sheets_col_dict)