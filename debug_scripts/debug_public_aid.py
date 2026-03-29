import pandas as pd
import json

try:
    xl = pd.ExcelFile('PathFinder input.xlsx')
    if 'PUBLIC AID' in xl.sheet_names:
        df = xl.parse('PUBLIC AID', header=None)
        # print the first 30 rows to see the structure
        print("--- PUBLIC AID SHEET ---")
        for i, row in df.head(30).iterrows():
            print([str(x) for x in row.values])
    else:
        print("Sheet 'PUBLIC AID' not found. Available sheets:", xl.sheet_names)
except Exception as e:
    print("Error:", e)
