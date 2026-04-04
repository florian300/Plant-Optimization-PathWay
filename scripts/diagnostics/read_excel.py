import pandas as pd
import sys

try:
    xl = pd.ExcelFile('PathFinder input.xlsx')
    print("Sheets:", xl.sheet_names)
    
    if 'OVERVIEW' in xl.sheet_names:
        df = xl.parse('OVERVIEW', header=None)
        print("--- OVERVIEW ---")
        for _, row in df.iterrows():
            vals = [str(x).strip() for x in row.values if pd.notna(x)]
            if vals:
                print(vals)
                
    if 'REFINERY' in xl.sheet_names:
        df = xl.parse('REFINERY', header=None)
        print("--- REFINERY ---")
        # print first few rows
        for _, row in df.head(40).iterrows():
            vals = [str(x).strip() for x in row.values if pd.notna(x)]
            if vals:
                print(vals)
                
except Exception as e:
    print(e)
