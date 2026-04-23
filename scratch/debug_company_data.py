
import sys
from pathlib import Path

# Add src to sys.path
repo_root = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2")
sys.path.insert(0, str(repo_root / "src"))

from pathway.core.ingestion import PathFinderParser

def debug_company_data():
    excel_path = repo_root / "data" / "raw" / "excel" / "PathFinder input.xlsx"
    print(f"Excel Path: {excel_path}")
    print(f"Exists: {excel_path.exists()}")
    
    parser = PathFinderParser(str(excel_path))
    data = parser.get_company_explorer_data()
    print(f"Company Data keys: {list(data.keys())}")
    for k, v in data.items():
        print(f"\nEntity: {k}")
        for block, rows in v.items():
            print(f"  {block}: {len(rows)} rows")

if __name__ == "__main__":
    debug_company_data()
