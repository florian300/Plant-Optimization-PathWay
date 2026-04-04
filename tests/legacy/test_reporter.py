import pandas as pd
import matplotlib.pyplot as plt
from reporting import PathFinderReporter

class DummyOpt:
    def __init__(self):
        self.years = list(range(2025, 2051))
        self.entity = type('obj', (object,), {'base_emissions': 1000000, 'sv_act_mode': 'PI'})
        self.data = type('obj', (object,), {
            'resources': ['CO2_EM', 'EN_ELEC'],
            'technologies': {},
            'objectives': [
                type('obj', (object,), {'resource':'CO2_EM', 'limit_type':'CAP', 'target_year':2035, 'cap_value':-0.5, 'comparison_year':2025}),
                type('obj', (object,), {'resource':'CO2_EM', 'limit_type':'CAP', 'target_year':2050, 'cap_value':-0.9, 'comparison_year':2025})
            ],
            'time_series': type('obj', (object,), {'carbon_quotas_pi': {}, 'carbon_quotas_norm': {}})
        })

df = pd.DataFrame({
    'Year': list(range(2025, 2051)),
    'Direct_CO2': [1000000]*26,
    'Indirect_CO2': [100000]*26,
    'Total_CO2': [1100000]*26,
    'Taxed_CO2': [800000]*26,
    'Free_Quota': [200000]*26
})

reporter = PathFinderReporter(DummyOpt())
reporter._plot_co2_trajectory(df)
print("Plot generated: CO2_Trajectory.png")
