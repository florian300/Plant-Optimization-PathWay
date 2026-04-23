
import pandas as pd
from pathlib import Path

def inspect_refinery_process():
    excel_path = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\data\raw\excel\PathFinder input.xlsx")
    df = pd.read_excel(excel_path, sheet_name='REFINERY', header=None)
    
    for i in range(40, 60):
        row = df.iloc[i]
        non_empty = [f"Col{j}: {val}" for j, val in enumerate(row) if pd.notna(val)]
        if non_empty:
            print(f"Row {i}: {' | '.join(non_empty)}")

if __name__ == "__main__":
    inspect_refinery_process()
