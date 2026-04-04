import pandas as pd, glob
from pathlib import Path
files = glob.glob('artifacts/reports/**/*.xlsx', recursive=True)
def load_transition_balance_table(workbook: Path) -> pd.DataFrame:
    try:
        raw = pd.read_excel(workbook, sheet_name='Charts', header=None)
    except ValueError:
        return pd.DataFrame()
    header_row = None
    for idx, row in raw.iterrows():
        values = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if any(v == 'Year' for v in values) and any('Net_Cumulative_Cost' in v for v in values):
            header_row = idx
            break
    if header_row is None:
        return pd.DataFrame()
    df = pd.read_excel(workbook, sheet_name='Charts', header=header_row)
    df = df.dropna(how='all')
    unnamed = [col for col in df.columns if str(col).startswith('Unnamed:')]
    if unnamed: df = df.drop(columns=unnamed)
    if 'Year' not in df.columns:
        if 'Net_Cumulative_Cost' in df.columns: df = df.rename(columns={df.columns[0]: 'Year'})
    return df.reset_index(drop=True)
print(load_transition_balance_table(Path(files[-1])).head())
