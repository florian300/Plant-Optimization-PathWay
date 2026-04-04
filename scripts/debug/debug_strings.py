import pandas as pd

file_path = 'PathFinder input.xlsx'
xl = pd.ExcelFile(file_path)

for sheet_name in ['RESSOURCES_PRICE', 'CARBON QUOTAS']:
    print(f"\n--- Strings in sheet: {sheet_name} ---")
    df = xl.parse(sheet_name, header=None)
    # Extract all non-numeric, non-null values
    all_values = df.values.flatten()
    strings = [str(v) for v in all_values if pd.notna(v) and not isinstance(v, (int, float))]
    
    unique_strings = sorted(list(set(strings)))
    for s in unique_strings:
        if any(keyword in s.upper() for keyword in ['LINEAR', 'BROWN', '&']):
            print(f"MATCH: '{s}'")
        else:
            # print first 20 chars of other strings to see what's there
            pass

print("\n--- Full search for '&' in all sheets ---")
for sheet_name in xl.sheet_names:
    df = xl.parse(sheet_name, header=None)
    mask = df.astype(str).str.contains('&', na=False)
    if mask.any().any():
        print(f"Found '&' in sheet '{sheet_name}'")
