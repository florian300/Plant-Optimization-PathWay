import pandas as pd
xl = pd.ExcelFile('PathFinder input.xlsx')

print("=== CARBON QUOTAS ===")
df = xl.parse('CARBON QUOTAS', header=None)
for i, row in df.iterrows():
    vals = [str(x).strip() for x in row.values if pd.notna(x)]
    if vals:
        print(f'Row {i}:', vals[:8])

print("\n=== NEW TECH (scenario refs) ===")
df = xl.parse('NEW TECH', header=None)
for i, row in df.iterrows():
    vals = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
    if vals and any(v.upper() in ['BS','CT','LCB','SCENARIO'] for v in vals):
        print(f'Row {i}:', vals[:10])

print("\n=== REFINERY (scenario refs) ===")
df = xl.parse('REFINERY', header=None)
for i, row in df.iterrows():
    vals = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip()]
    if vals and any(v.upper() in ['BS','CT','LCB','SCENARIOS','SCENARIO'] for v in vals):
        print(f'Row {i}:', vals[:10])
