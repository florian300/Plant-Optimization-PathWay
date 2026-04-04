import pandas as pd, glob
files = glob.glob('artifacts/reports/**/*.xlsx', recursive=True)
for f in files:
    try:
        df = pd.read_excel(f, sheet_name='Charts', header=None)
        print(f, 'has Charts, rows:', len(df))
    except ValueError:
        print(f, 'NO Charts')
