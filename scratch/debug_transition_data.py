
import sys
from pathlib import Path

# Add src to sys.path
repo_root = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2")
sys.path.insert(0, str(repo_root / "src"))

from pathway.core.ingestion import PathFinderParser

def debug_transition_data():
    excel_path = repo_root / "data" / "raw" / "excel" / "PathFinder input.xlsx"
    parser = PathFinderParser(str(excel_path))
    data = parser.get_company_explorer_data()
    
    if 'TOTALENERGIES' in data:
        te = data['TOTALENERGIES']
        if 'TRANSITION' in te:
            print(f"TRANSITION Rows: {len(te['TRANSITION'])}")
            for row in te['TRANSITION']:
                print(row)
        else:
            print("TRANSITION block missing")
    else:
        print("TOTALENERGIES not found")

if __name__ == "__main__":
    debug_transition_data()
