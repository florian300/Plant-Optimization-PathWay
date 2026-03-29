import pandas as pd
import json

file_path = 'PathFinder input.xlsx'

xl = pd.ExcelFile(file_path)
info = {}

for sheet_name in xl.sheet_names:
    df = xl.parse(sheet_name, header=None)
    # Find the row that might contain the actual headers
    # A heuristic: row with the most non-null string values
    non_null_counts = df.notna().sum(axis=1)
    # let's just dump the first 20 rows, converting to dict
    head_data = df.head(15).fillna("").values.tolist()
    info[sheet_name] = head_data

with open('excel_structure.json', 'w', encoding='utf-8') as f:
    json.dump(info, f, indent=2, ensure_ascii=False)
    
print("Dumped structure to excel_structure.json")
