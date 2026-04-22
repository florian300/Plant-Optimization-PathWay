import sys
import os
import json
import pandas as pd
from unittest.mock import MagicMock

# Add src to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from pathway.core.reporting import PathFinderReporter
from pathway.core.model import PathFinderData, Resource

def verify():
    # Load the junk results
    results_path = 'c:/Users/flori/Documents/TFM/PathWay_Python_Tool - v2/artifacts/reports/TOTALENERGIES/CLEAN TECH/CLEAN TECH_results.json'
    if not os.path.exists(results_path):
        print("Results file not found. Run the simulation first or provide a valid path.")
        return

    with open(results_path, 'r') as f:
        results = json.load(f)

    # Mock Optimizer
    mock_opt = MagicMock()
    mock_opt.years = [int(y) for y in results['years']]
    mock_opt.entity.id = "Base Company"
    mock_opt.entity.name = "Base Company"
    
    # Mock resources with categories
    resources = {
        'W1': Resource(id='W1', name='WATER', type='WATER', unit='m3', category='WATER', resource_type='WATER'),
        'EN_ELEC': Resource(id='EN_ELEC', name='ELECTRICITY', type='ENERGY', unit='MWh', category='ENERGY', resource_type='ELECTRICITY'),
        'EN_FUEL': Resource(id='EN_FUEL', name='FUEL', type='ENERGY', unit='MWh', category='ENERGY', resource_type='FUEL')
    }
    
    data = MagicMock()
    data.resources = resources
    data.time_series.resource_prices = {r: {y: 10.0 for y in mock_opt.years} for r in resources}
    data.time_series.other_emissions_factors = {}
    data.reporting_toggles.chart_transition_costs = True
    data.reporting_toggles.results_excel = False
    
    mock_opt.data = data
    
    # Mock variables from results
    # We need to populate mock_opt.cons_vars, mock_opt.invest_vars, etc.
    mock_opt.cons_vars = {}
    for y_str, res_vals in results['consumption'].items():
        y = int(y_str)
        for r_id, val in res_vals.items():
            var = MagicMock()
            var.varValue = val
            mock_opt.cons_vars[(y, r_id)] = var
            
    mock_opt.invest_vars = {}
    for y_str, tech_vals in results['investments'].items():
        y = int(y_str)
        for p_tech, val in tech_vals.items():
            if '##' in p_tech:
                p_id, t_id = p_tech.split('##')
                var = MagicMock()
                var.varValue = val
                mock_opt.invest_vars[(y, p_id, t_id)] = var

    # Reporter
    reporter = PathFinderReporter(mock_opt, scenario_id="CLEAN TECH", scenario_name="Clean Tech")
    reporter.verbose = True
    
    # Check energy drivers
    energy_ids = reporter._get_energy_resource_ids()
    print(f"Energy IDs: {energy_ids}")
    if 'W1' in energy_ids:
        print("FAIL: W1 should not be an energy ID")
    else:
        print("SUCCESS: W1 excluded from energy drivers")

    # Check safe var value
    junk_var = MagicMock()
    junk_var.varValue = 3.49e11
    safe_val = reporter._get_safe_var_value(junk_var)
    print(f"Junk value 3.49e11 -> Safe value: {safe_val}")
    if safe_val == 0.0:
        print("SUCCESS: Junk value capped at 0.0")
    else:
        print("FAIL: Junk value not capped")

if __name__ == "__main__":
    verify()
