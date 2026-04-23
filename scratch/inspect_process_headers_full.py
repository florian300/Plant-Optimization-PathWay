
import pandas as pd
from pathlib import Path

def inspect_process_headers_full():
    excel_path = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\data\raw\excel\PathFinder input.xlsx")
    df = pd.read_excel(excel_path, sheet_name='REFINERY', header=None)
    
    row = df.iloc[39]
    print(f"Row 39 Columns: {len(row)}")
    for j, val in enumerate(row):
        if pd.notna(val):
            print(f"Col {j}: {val}")

if __name__ == "__main__":
    inspect_process_headers_full()
