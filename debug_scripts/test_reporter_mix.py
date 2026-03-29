import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from reporting import PathFinderReporter

class Resource:
    def __init__(self, id, unit):
        self.id = id
        self.unit = unit

class DummyData:
    def __init__(self):
        self.resources = {
            'CO2_EM': Resource('CO2_EM', 'tCO2'),
            'EN_ELEC': Resource('EN_ELEC', 'MWH'),
            'EN_GAS': Resource('EN_GAS', 'MWH'),
            'EN_LGP': Resource('EN_LGP', 't'),
            'EN_GASO': Resource('EN_GASO', 't'),
        }
        self.objectives = []
        
class DummyOpt:
    def __init__(self):
        self.years = list(range(2025, 2051))
        self.entity = type('obj', (object,), {'base_emissions': 1000000, 'sv_act_mode': 'PI'})
        self.data = DummyData()

df_cons = pd.DataFrame({
    'Year': list(range(2025, 2051)),
    'EN_ELEC': np.linspace(100, 500, 26),      # Consumption
    'EN_GAS': np.linspace(300, 100, 26),       # Consumption
    'EN_LGP': np.linspace(-50, -100, 26),      # Production
    'EN_GASO': np.linspace(-200, -250, 26),    # Production
    'CO2_EM': np.linspace(0, 0, 26)            # Ignored
})

reporter = PathFinderReporter(DummyOpt())
reporter._plot_energy_mix(df_cons)
print("Plots generated in Results/ directory.")
