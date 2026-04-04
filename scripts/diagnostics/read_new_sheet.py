import pandas as pd

xl = pd.ExcelFile('PathFinder input.xlsx')
if 'NEW TECH_INDIRECT' in xl.sheet_names:
    df = xl.parse('NEW TECH_INDIRECT', header=None)
    for _, row in df.iterrows():
        vals = [str(x) for x in row.values if pd.notna(x)]
        if vals:
            print(vals)
