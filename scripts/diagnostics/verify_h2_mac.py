
import os
import sys
import pandas as pd
from pathway.core.ingestion import PathFinderParser
from pathway.core.optimizer import PathFinderOptimizer
from pathway.core.reporting import PathFinderReporter

def test_h2_mac():
    print("Starting H2 MAC verification...")
    # 1. Parse Data
    parser = PathFinderParser("PathFinder input.xlsx")
    data = parser.parse()
    
    # 2. Run Optimization (Short duration to save time)
    data.parameters.duration = 5 
    optimizer = PathFinderOptimizer(data, verbose=True)
    optimizer.build_model()
    optimizer.solve()
    
    # 3. Generate Report
    # We use a custom results dir
    reporter = PathFinderReporter(optimizer, scenario_name="H2_MAC_Verify", verbose=True)
    reporter.generate_report()
    
    # 4. Verify Results
    results_file = os.path.join("Results", "H2_MAC_Verify", "CO2_Abatement_Cost.png")
    if os.path.exists(results_file):
        print(f"SUCCESS: MAC chart generated at {results_file}")
    else:
        print("FAILURE: MAC chart not found.")
        
    # Also check the Excel data if possible to see the MAC value
    excel_file = os.path.join("Results", "H2_MAC_Verify", f"PathFinder_Report_H2_MAC_Verify.xlsx")
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file, sheet_name="MAC_Data")
        print("MAC Data from Excel:")
        print(df[['Project', 'MAC (€/tCO2)', 'Status']])
        
        # Check if FUEL_TO_H2 (labeled as 'CHANGE FUEL FOR H2' or similar) has a valid MAC
        h2_techs = df[df['Project'].str.contains('H2', case=False)]
        if not h2_techs.empty:
            print("Found Hydrogen technologies in MAC data.")
        else:
            print("WARNING: No Hydrogen technologies found in MAC data.")
    else:
        print("Excel report not found.")

if __name__ == "__main__":
    test_h2_mac()
