
import sys
from pathlib import Path

# Add src to sys.path
repo_root = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2")
sys.path.insert(0, str(repo_root / "src"))

from pathway.core.ingestion import PathFinderParser

def debug_total_keys():
    excel_path = repo_root / "data" / "raw" / "excel" / "PathFinder input.xlsx"
    parser = PathFinderParser(str(excel_path))
    data = parser.get_company_explorer_data()
    
    if 'TOTALENERGIES' in data:
        te = data['TOTALENERGIES']
        if 'TOTAL' in te and te['TOTAL']:
            first_row = te['TOTAL'][0]
            print(f"Keys: {list(first_row.keys())}")
            print(f"First row: {first_row}")
        else:
            print("TOTAL block empty or missing for TOTALENERGIES")
    else:
        print("TOTALENERGIES not found in data")

if __name__ == "__main__":
    debug_total_keys()
