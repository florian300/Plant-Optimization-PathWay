import pandas as pd
import sys

try:
    xl = pd.ExcelFile('PathFinder input.xlsx')
    df_oe = xl.parse('OTHER EMISSIONS', header=None)
    for _, row in df_oe.iterrows():
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
                    em_val = row.values[i+1]
                    unit = str(row.values[i+2]) if i+2 < len(row.values) else ''
                    if 'H2' in r_id or r_id == 'EN_ELEC':
                        print(f"Year {year} | Res: {r_id} | Factor: {em_val} | Unit: {unit}")
except Exception as e:
    print(e)
