import pandas as pd
import matplotlib.pyplot as plt
import os
from reporting import PathFinderReporter

class DummyOpt:
    def __init__(self):
        self.years = list(range(2025, 2031))
        # Mock entity
        self.entity = type('obj', (object,), {
            'id': 'ENT1',
            'base_emissions': 1000000,
            'base_consumptions': {'EN_ELEC': 500000, 'EN_FUEL': 200000},
            'sv_act_mode': 'PI',
            'annual_operating_hours': 8760,
            'ca_percentage_limit': 0.05,
            'sold_resources': []
        })
        
        # Mock resources
        res_elec = type('obj', (object,), {'id': 'EN_ELEC', 'name': 'Electricity', 'unit': 'MWh'})
        res_fuel = type('obj', (object,), {'id': 'EN_FUEL', 'name': 'Natural Gas', 'unit': 'GJ'})
        res_co2 = type('obj', (object,), {'id': 'CO2_EM', 'name': 'CO2 Emissions', 'unit': 'tCO2'})
        
        # Mock data
        self.data = type('obj', (object,), {
            'resources': {'EN_ELEC': res_elec, 'EN_FUEL': res_fuel, 'CO2_EM': res_co2},
            'technologies': {},
            'objectives': [],
            'time_series': type('obj', (object,), {
                'carbon_quotas_pi': {y: 0.5 for y in self.years},
                'carbon_quotas_norm': {y: 0.5 for y in self.years},
                'carbon_prices': {y: 100.0 for y in self.years},
                'other_emissions_factors': {
                    'EN_ELEC': {y: 0.2 for y in self.years},
                    'EN_FUEL': {y: 0.05 for y in self.years}
                },
                'resource_prices': {}
            })
        })
        
        # Mock solver variables
        self.cons_vars = {(y, r): type('obj', (object,), {'varValue': 100000 if r=='EN_ELEC' else 50000}) for y in self.years for r in ['EN_ELEC', 'EN_FUEL']}
        self.emis_vars = {y: type('obj', (object,), {'varValue': 800000}) for y in self.years}
        self.taxed_emis_vars = {y: type('obj', (object,), {'varValue': 300000}) for y in self.years}
        self.active_vars = {}
        self.invest_vars = {}
        self.penalty_vars = []
        self.model = type('obj', (object,), {'variablesDict': lambda: {}})

# Run the test
print("Starting verification test...")
opt = DummyOpt()
reporter = PathFinderReporter(opt)

# Mock some internal calls to avoid full report generation complexity
df_indir = []
for t in opt.years:
    row = {'Year': t}
    for r_id in opt.data.resources:
        if r_id in opt.data.time_series.other_emissions_factors:
            factor = opt.data.time_series.other_emissions_factors[r_id].get(t, 0.0)
            if factor > 0:
                cons_val = opt.cons_vars[(t, r_id)].varValue
                row[r_id] = cons_val * factor
    df_indir.append(row)
df_indir = pd.DataFrame(df_indir)

print("Plotting indirect emissions...")
reporter._plot_indirect_emissions(df_indir)

if os.path.exists('Results/Indirect_Emissions.png'):
    print("SUCCESS: Indirect_Emissions.png generated.")
else:
    print("FAILURE: Indirect_Emissions.png NOT generated.")

print("Verification test finished.")
