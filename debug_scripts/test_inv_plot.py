import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from reporting import PathFinderReporter

class DummyEntity:
    def __init__(self):
        self.ca_percentage_limit = 0.05
        self.processes = {
            'PROC_A': type('obj', (object,), {'name': 'Process Alpha'}),
            'PROC_B': type('obj', (object,), {'name': 'Process Beta'})
        }

class DummyOpt:
    def __init__(self):
        self.years = list(range(2025, 2030))
        self.entity = DummyEntity()
        self.data = type('obj', (object,), {
            'resources': {},
            'technologies': {
                'TECH_1': type('obj', (object,), {'name': 'Tech 1 Name'}),
                'TECH_2': type('obj', (object,), {'name': 'Tech 2 Name'}),
            },
            'objectives': [],
            'time_series': type('obj', (object,), {'carbon_quotas_pi': {}, 'carbon_quotas_norm': {}})
        })

reporter = PathFinderReporter(DummyOpt())
reporter.df_costs = pd.DataFrame({
    'Year': [2025, 2026, 2027, 2028, 2029],
    'Budget_Limit': [50_000_000]*5
})
reporter.df_projects = pd.DataFrame({
    'Year': [2025, 2026, 2027, 2028, 2029],
    'PROC_A##TECH_1': [10_000_000, 15_000_000, 0, 0, 0],
    'PROC_A##TECH_2': [0, 0, 20_000_000, 10_000_000, 0],
    'PROC_B##TECH_3': [5_000_000, 5_000_000, 5_000_000, 5_000_000, 5_000_000],
    'Financing Interests': [0, 1_000_000, 2_000_000, 1_000_000, 500_000]
})

reporter._plot_investment_costs(pd.DataFrame())
print("Plot generated: Investment_Costs.png")
