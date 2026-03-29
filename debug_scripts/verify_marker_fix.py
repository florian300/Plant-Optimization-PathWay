import pandas as pd
import matplotlib.pyplot as plt
import sys
import os

# Add the project root to sys.path to import core modules
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from core.reporting import PathFinderReporter

class DummyData:
    def __init__(self):
        self.resources = {'CO2_EM': type('obj', (object,), {'id': 'CO2_EM', 'name': 'CO2 Emissions', 'unit': 'tCO2'})}
        self.technologies = {}
        self.objectives = [
            type('obj', (object,), {
                'resource': 'CO2_EM', 
                'target_year': 2050, 
                'cap_value': 0.0, 
                'comparison_year': None,
                'name': 'Net Zero Goal',
                'group': 'Main',
                'mode': 'POINT'
            })
        ]

class DummyOpt:
    def __init__(self):
        self.years = list(range(2025, 2051))
        self.entity = type('obj', (object,), {'base_emissions': 1000000})
        self.data = DummyData()
        self.taxed_emis_vars = {t: type('obj', (object,), {'varValue': 0.0}) for t in self.years}
        self.emis_vars = {t: type('obj', (object,), {'varValue': 0.0}) for t in self.years}

df = pd.DataFrame({
    'Year': list(range(2025, 2051)),
    'Direct_CO2': [1000000]*26,
    'Indirect_CO2': [200000]*26,
    'Total_CO2': [1200000]*26, # Old total
    'Taxed_CO2': [0]*26,
    'Free_Quota': [0]*26,
    'DAC_Captured_kt': [100]*26, # 100 ktCO2
    'Credits_Purchased_kt': [50]*26 # 50 ktCO2
})

reporter = PathFinderReporter(DummyOpt())
# Mock _apply_premium_style and _add_watermark to avoid issues with missing fonts/images
reporter._apply_premium_style = lambda ax: None
reporter._add_watermark = lambda fig: None

# Call the plot method
reporter._plot_co2_trajectory(df)
print("Verification plot generated: Results/CO2_Trajectory.png")
print("Expected Net Direct: 1000 - 100 - 50 = 850 ktCO2")
print("Expected Total (New): 850 + 200 = 1050 ktCO2")
print("Please check if the red dashed line is at 1050 ktCO2.")
