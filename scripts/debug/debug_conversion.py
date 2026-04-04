import pandas as pd
xl = pd.ExcelFile('PathFinder input.xlsx')
df = xl.parse('OverView', header=None)

print("--- UNIT CONVERSION ---")
start_idx = -1
for i, row in df.iterrows():
    if "UNIT CONVERSION" in [str(x).strip().upper() for x in row.values if pd.notna(x)]:
        start_idx = i
        break

if start_idx != -1:
    for i, row in df.iloc[start_idx:start_idx+20].iterrows():
        vals = [x for x in row if pd.notna(x) and str(x).strip() != '']
        print(f"Row {i}: {vals}")
