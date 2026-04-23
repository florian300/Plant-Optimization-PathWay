
import sys
from pathlib import Path

# Add src to sys.path
repo_root = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2")
sys.path.insert(0, str(repo_root / "src"))

from pathway.core.ingestion import PathFinderParser

def debug_process_keys():
    excel_path = repo_root / "data" / "raw" / "excel" / "PathFinder input.xlsx"
    parser = PathFinderParser(str(excel_path))
    data = parser.get_company_explorer_data()
    
    if 'TOTALENERGIES' in data:
        te = data['TOTALENERGIES']
        if 'PROCESS' in te and te['PROCESS']:
            print(f"Keys: {list(te['PROCESS'][0].keys())}")
            for p in te['PROCESS']:
                print(f"Row: {p.get('PROCESS NAME')} | {p.get('ID')}")
        else:
            print("PROCESS block empty or missing")
    else:
        print("TOTALENERGIES not found in data")

if __name__ == "__main__":
    debug_process_keys()
