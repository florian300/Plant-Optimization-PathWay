import pandas as pd
import os

file_path = r'c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\PathFinder input.xlsx'
xl = pd.ExcelFile(file_path)
df_prices = xl.parse('RESSOURCES_PRICE', header=None)

raw_prices = {}
for _, row in df_prices.iterrows():
    year = None
    start_idx = -1
    for idx, cell in enumerate(row.values):
        try:
            y = int(cell)
            if y >= 2000 and y <= 2100:
                year = y
                start_idx = idx
                break
        except: pass
    
    if year is not None and start_idx != -1:
        for i in range(start_idx + 1, len(row.values) - 1, 3):
            r_id = str(row.values[i]).strip()
            if r_id and r_id != 'nan' and pd.notna(row.values[i+1]):
                if r_id not in raw_prices:
                    raw_prices[r_id] = []
                raw_prices[r_id].append(year)

print(f"Parsed {len(raw_prices)} resources from PRICES sheet:")
for r_id in sorted(raw_prices.keys()):
    print(f" - {r_id}")
