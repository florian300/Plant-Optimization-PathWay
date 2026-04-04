import pandas as pd
import sys

xl = pd.ExcelFile('PathFinder input.xlsx')

for sheet in ['RESSOURCES_PRICE', 'CARBON QUOTAS']:
    df = xl.parse(sheet)
    for idx, row in df.head(20).iterrows():
        # check if there's a year
        if 2025 in row.values or '2025' in row.values or 2025.0 in row.values:
            print(f"Sheet {sheet} Year row {idx}: {row.tolist()[:5]}")
            df.columns = [str(c).strip() if isinstance(c, str) else c for c in row]
            break
            
    print(f"Sheet {sheet} cols: {df.columns.tolist()[:5]}")
    print(df.head(5).to_string())
    print("-------")
