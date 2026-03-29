import pandas as pd

file_path = 'PathFinder input.xlsx'
xl = pd.ExcelFile(file_path)

found = False
for sheet in xl.sheet_names:
    df = xl.parse(sheet, header=None)
    for r_idx, row in df.iterrows():
        for c_idx, val in enumerate(row):
            if pd.notna(val) and isinstance(val, str):
                if 'BROWNIEN' in val.upper():
                    print(f"FOUND: '{val}' in sheet '{sheet}', cell ({r_idx+1}, {c_idx+1})")
                    found = True
                if 'LINEAR&' in val.upper():
                    print(f"FOUND: '{val}' in sheet '{sheet}', cell ({r_idx+1}, {c_idx+1})")
                    found = True

if not found:
    print("Absolutely NO 'BROWNIEN' or 'LINEAR&' keywords found in the entire Excel file.")
    # Debug: show some headers
    print("\nSheet names:", xl.sheet_names)
