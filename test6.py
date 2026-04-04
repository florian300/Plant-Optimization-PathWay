import pandas as pd, glob
files = glob.glob('artifacts/reports/**/*.xlsx', recursive=True)
df = pd.read_excel(files[-1], sheet_name='Charts', header=None)
for i, row in df.iterrows():
    for v in row.values:
        if isinstance(v, str) and 'Ecological' in v:
            print('Line:', i, v)
