import pandas as pd

xl = pd.ExcelFile('PathFinder input.xlsx')
df = xl.parse('NEW TECH', header=None)

print("--- NEW TECH SHEET ---")
for i, row in df.head(50).iterrows():
    vals = [x for x in row if pd.notna(x) and str(x).strip() != '']
    print(f"Row {i}: {vals}")
