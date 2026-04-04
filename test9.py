import pandas as pd, glob
files = glob.glob('artifacts/reports/**/*.xlsx', recursive=True)
df = pd.read_excel(files[-1], sheet_name='Charts', header=None)
print('Length of charts:', len(df))
