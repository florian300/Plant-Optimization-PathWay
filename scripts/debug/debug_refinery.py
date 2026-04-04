import pandas as pd

xl = pd.ExcelFile('PathFinder input.xlsx')
df = xl.parse('REFINERY', header=None)

print("--- REFINERY TOTAL BLOCK ---")
for i, row in df.iloc[14:35].iterrows():
    vals = [x for x in row.tolist() if pd.notna(x) and str(x).strip() != '']
    print(f"Row {i}: {vals}")
