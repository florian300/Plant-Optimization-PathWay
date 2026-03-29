import pandas as pd
import sys

file_path = 'PathFinder input.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    print("Sheets found:", xl.sheet_names)
    print("="*50)
    for sheet_name in xl.sheet_names:
        print(f"\n--- Sheet: {sheet_name} ---")
        df = xl.parse(sheet_name)
        print("Columns:", df.columns.tolist())
        print("First 3 rows:")
        print(df.head(3).to_string())
except Exception as e:
    print("Error:", e)
sys.exit(0)
