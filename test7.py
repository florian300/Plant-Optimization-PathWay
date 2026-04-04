import pandas as pd, glob
files = glob.glob('artifacts/reports/**/*.xlsx', recursive=True)
df = pd.read_excel(files[-1], sheet_name='Charts', header=None)
for i, row in df.iterrows():
    if row.notna().sum() == 1:
        print('Title:', row.dropna().iloc[0])
