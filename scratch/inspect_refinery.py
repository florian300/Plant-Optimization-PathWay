
import pandas as pd
from pathlib import Path

def inspect_refinery():
    excel_path = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\data\raw\excel\PathFinder input.xlsx")
    df = pd.read_excel(excel_path, sheet_name='REFINERY', header=None)
    
    for i, row in df.iterrows():
        non_empty = [f"Col{j}: {val}" for j, val in enumerate(row) if pd.notna(val)]
        if non_empty:
            # Check for any of the markers
            row_str = " ".join([str(v) for v in row if pd.notna(v)]).upper()
            markers = ['REF', 'TOTAL', 'PROCESS', 'TECHNOLOGICAL TRANSITION', 'START', 'END']
            if any(m in row_str for m in markers):
                print(f"Row {i}: {' | '.join(non_empty)}")

if __name__ == "__main__":
    inspect_refinery()
