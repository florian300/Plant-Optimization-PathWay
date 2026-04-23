
import pandas as pd
from pathlib import Path

def search_markers():
    excel_path = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\data\raw\excel\PathFinder input.xlsx")
    xl = pd.ExcelFile(excel_path)
    markers = ['INIT', 'REF', 'TOTAL', 'PROCESS', 'TECHNOLOGICAL TRANSITION']
    
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)
        for i, row in df.iterrows():
            row_str = " ".join([str(v) for v in row if pd.notna(v)]).upper()
            for m in markers:
                if m in row_str:
                    print(f"Sheet: {sheet_name} | Row: {i} | Marker: {m} | Content: {row_str}")

if __name__ == "__main__":
    search_markers()
