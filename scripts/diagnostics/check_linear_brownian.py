import pandas as pd
import numpy as np

file_path = 'PathFinder input.xlsx'
xl = pd.ExcelFile(file_path)

found = []

for sheet_name in xl.sheet_names:
    df = xl.parse(sheet_name, header=None)
    # Convert all values to string and search for "LINEAR&BROWNIEN" or subset
    # Let's search for "&" specifically to see if that's being missed
    mask1 = df.astype(str).apply(lambda x: x.str.contains('LINEAR&BROWNIEN', case=False, na=False))
    mask2 = df.astype(str).apply(lambda x: x.str.contains('LINEAR & BROWNIEN', case=False, na=False))
    mask3 = df.astype(str).apply(lambda x: x.str.contains('BROWNIEN', case=False, na=False))
    
    mask = mask1 | mask2 | mask3
    
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
    print(f"Found {len(found)} occurrences:")
    for f in found:
        print(f"- Sheet '{f['sheet']}', Row {f['row']}, Col {f['col']}: '{f['value']}'")
else:
    print("No occurrences of 'LINEAR&BROWNIEN' or 'BROWNIEN' found.")
    # Show some non-null values for debugging
    print("\nSample values from RESSOURCES_PRICE sheet:")
    df_price = xl.parse('RESSOURCES_PRICE', header=None)
    print(df_price.head(50).iloc[:, 0:10]) # show first 50 rows, first 10 cols
