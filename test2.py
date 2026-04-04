import pandas as pd, sys, glob
files = glob.glob('artifacts/reports/**/Master_Plan.xlsx', recursive=True)
if not files: sys.exit(0)
df = pd.read_excel(files[0], sheet_name='Charts')
for i, row in df.iterrows():
    if any(isinstance(v, str) and 'Ecological' in v for v in row.values):
        print('Found at', i)
        print(df.iloc[i:i+10].values)
        break
