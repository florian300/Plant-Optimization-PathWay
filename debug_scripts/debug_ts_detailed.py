import pandas as pd

xl = pd.ExcelFile('PathFinder input.xlsx')
print(f"--- OTHER EMISSIONS ---")
df = xl.parse('OTHER EMISSIONS', header=None)
for i, row in df.head(30).iterrows():
    vals = [x for x in row.tolist() if pd.notna(x) and str(x).strip() != '']
    print(f"Row {i}: {vals}")
print("\n")
