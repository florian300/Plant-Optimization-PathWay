import pandas as pd

file_path = 'PathFinder input.xlsx'
xl = pd.ExcelFile(file_path)

for sheet in xl.sheet_names:
    df = xl.parse(sheet, header=None)
    for r in range(len(df)):
        for c in range(len(df.columns)):
            val = df.iloc[r, c]
            if pd.notna(val):
                s_val = str(val).strip().upper()
                if 'LINEAR&BROWNIEN' in s_val:
                    print(f"BINGO! Found '{val}' in sheet '{sheet}', cell ({r+1}, {c+1})")
                elif 'LINEAR' in s_val and '&' in s_val and 'BROWN' in s_val:
                    print(f"ALMOST BINGO! Found '{val}' in sheet '{sheet}', cell ({r+1}, {c+1})")
                elif '&' in s_val:
                    # just print any string with & for debugging
                    if len(s_val) < 50:
                        print(f"AMBIGUOUS: '{val}' in sheet '{sheet}', cell ({r+1}, {c+1})")

print("\n--- End of refined search ---")
