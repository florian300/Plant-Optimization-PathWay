
import pandas as pd
from pathlib import Path

def inspect_overview():
    excel_path = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\data\raw\excel\PathFinder input.xlsx")
    df = pd.read_excel(excel_path, sheet_name='OverView')
    
    print("Columns:", df.columns.tolist())
    print("\nFirst 30 rows:")
    print(df.head(30))
    
    # Search for blocks
    markers = ['INIT', 'REF', 'TOTAL', 'PROCESS', 'TECHNOLOGICAL TRANSITION']
    for marker in markers:
        matches = df[df.apply(lambda row: row.astype(str).str.contains(marker).any(), axis=1)]
        if not matches.empty:
            print(f"\nFound marker '{marker}' at rows: {matches.index.tolist()}")

if __name__ == "__main__":
    inspect_overview()
