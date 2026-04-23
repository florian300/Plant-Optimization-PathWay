
import pandas as pd
from pathlib import Path

def detailed_inspect_overview():
    excel_path = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\data\raw\excel\PathFinder input.xlsx")
    df = pd.read_excel(excel_path, sheet_name='OverView', header=None)
    
    for i, row in df.iterrows():
        # Filter out rows that are completely empty
        non_empty = [f"Col{j}: {val}" for j, val in enumerate(row) if pd.notna(val)]
        if non_empty:
            print(f"Row {i}: {' | '.join(non_empty)}")

if __name__ == "__main__":
    detailed_inspect_overview()
