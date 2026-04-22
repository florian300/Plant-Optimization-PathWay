import sys
import os
import pandas as pd
from unittest.mock import MagicMock

# Add src to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from pathway.core.optimizer import PathFinderOptimizer
from pathway.core.model import PathFinderData, Resource, TimeSeriesData

def verify():
    # Mock data
    mock_data = MagicMock()
    mock_data.parameters = MagicMock()
    mock_data.parameters.start_year = 2025
    mock_data.parameters.duration = 5
    mock_data.parameters.discount_rate = 0.0
    mock_data.parameters.time_limit = 10
    mock_data.parameters.mip_gap = 0.1
    mock_data.dac_params.active = False
    mock_data.credit_params.active = False
    mock_data.ccfd_params.active = False
    mock_data.grant_params.active = False
    
    # Resource with Indirect Tax
    res_h2 = Resource(id='EN_H2', type='CONSUMPTION', unit='GJ', name='Hydrogen', tax_indirect_emissions=True, resource_type='HYDROGEN')
    res_co2 = Resource(id='CO2_EM', type='EMISSIONS', unit='tCO2', name='CO2', resource_type='CO2')
    mock_data.resources = {'EN_H2': res_h2, 'CO2_EM': res_co2}
    
    # Time Series
    ts = MagicMock(spec=TimeSeriesData)
    ts.carbon_prices = {2025: 100.0, 2026: 100.0, 2027: 100.0, 2028: 100.0, 2029: 100.0}
    ts.carbon_penalties = {y: 0.0 for y in range(2025, 2030)}
    ts.carbon_quotas_norm = {y: 1.0 for y in range(2025, 2030)}
    ts.other_emissions_factors = {'EN_H2': {y: 0.05 for y in range(2025, 2030)}} # 50 kgCO2 / GJ
    ts.resource_limits = {}
    ts.resource_prices = {}
    mock_data.time_series = ts
    
    # Entity
    entity = MagicMock()
    entity.id = "TOTAL"
    entity.start_year = 2025
    entity.end_year = 2050
    entity.ca_percentage_limit = 0
    entity.base_emissions = 1000.0
    entity.base_consumptions = {'EN_H2': 100.0}
    entity.processes = {}
    mock_data.entities = {"TOTAL": entity}
    
    # Build Optimizer
    opt = PathFinderOptimizer(mock_data, verbose=True)
    opt.build_model()
    
    # Check if objective function contains the indirect tax term for EN_H2
    # Tax = Consumption * 0.05 * 100 = Consumption * 5
    # So the coefficient for cons_vars[(2025, 'EN_H2')] in the objective should be 5.
    
    obj = opt.model.objective
    # obj is a linear expression. We check the coefficients.
    found = False
    for var, coeff in obj.items():
        # print(f"Var: {var.name}, Coeff: {coeff}")
        if f"cons_2025_EN_H2" in var.name or "EN_H2" in var.name:
            print(f"Candidate: {var.name}, Coeff: {coeff}")
        if f"cons_2025_EN_H2" in var.name.lower():
            print(f"Found variable: {var.name}, Coefficient: {coeff}")
            if abs(coeff - 5.0) < 1e-6:
                print("SUCCESS: Indirect tax coefficient is correct (5.0)")
                found = True
                break
    
    if not found:
        print("FAIL: Indirect tax term not found in objective function.")

if __name__ == "__main__":
    verify()
