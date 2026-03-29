import pandas as pd
import numpy as np

file_path = 'PathFinder input.xlsx'
xl = pd.ExcelFile(file_path)

found = []

for sheet_name in xl.sheet_names:
    df = xl.parse(sheet_name, header=None)
    # Search for "BROWNIEN" in all cells
    mask = df.astype(str).apply(lambda x: x.str.contains('BROWNIEN', case=False, na=False))
    if mask.any().any():
        # Find exactly where
        rows, cols = np.where(mask)
        for r, c in zip(rows, cols):
            val = df.iloc[r, c]
            found.append({
                'sheet': sheet_name,
                'row': r + 1,
                'col': c + 1,
                'value': val
            })

if found:
    print(f"Found {len(found)} occurrences of 'BROWNIEN':")
    for f in found:
        print(f"- Sheet '{f['sheet']}', Row {f['row']}, Col {f['col']}: '{f['value']}'")
else:
    print("No occurrences of 'BROWNIEN' found in the Excel file.")
