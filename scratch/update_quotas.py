
import sys
from pathlib import Path
import os

# Add src to sys.path
repo_root = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2")
sys.path.insert(0, str(repo_root / "src"))

from pathway.core.ingestion import PathFinderParser
from pathway.core.optimizer import PathFinderOptimizer
from pathway.core.reporting import PathFinderReporter

def update_charts():
    excel_path = repo_root / "data" / "raw" / "excel" / "PathFinder input.xlsx"
    parser = PathFinderParser(str(excel_path))
    # Parse for CLEAN TECH
    data = parser.parse(scenario_id="CT") # CT is the ID for Clean Tech
    
    # We need an optimizer object for the reporter
    optimizer = PathFinderOptimizer(data)
    
    reporter = PathFinderReporter(
        optimizer,
        scenario_id="CT",
        scenario_name="CLEAN TECH"
    )
    # Manually call the quotas plot
    reporter._plot_simulation_quotas(show_png=False)
    print("Updated Simulation_Quotas.json for CLEAN TECH")

if __name__ == "__main__":
    update_charts()
    # Now regenerate dashboard
    os.system(f"set PYTHONPATH=src && python \"{repo_root}/scripts/ops/generate_results_dashboard.py\"")
