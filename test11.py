import pandas as pd, glob
files = glob.glob('artifacts/reports/**/*.xlsx', recursive=True)
df = pd.read_excel(files[-1], sheet_name='Charts', header=None)
for i, r in df.iterrows():
    v = [str(x) for x in r if pd.notna(x)]
    if 'Net_Cumulative_Cost' in v: print(i, v)
