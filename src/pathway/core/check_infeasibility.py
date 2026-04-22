import pickle
import os
from pathway.core.optimizer import PathFinderOptimizer

# Try to find a pickled version of the data if available, or just re-parse
# Actually, I'll just use the ingestion module to parse the Excel and see the baseline.
from pathway.core.ingestion import PathFinderParser

excel_path = 'c:/Users/flori/Documents/TFM/PathWay_Python_Tool - v2/data/raw/excel/PathFinder input.xlsx'
parser = PathFinderParser(excel_path)
data = parser.parse()

for entity_id, entity in data.entities.items():
    print(f"Entity: {entity_id}")
    w1_cons = entity.base_consumptions.get('W1', 0.0)
    print(f"  W1 Baseline Consumption: {w1_cons:,.2f}")
    
    if 'W1' in data.time_series.resource_limits:
        limit = data.time_series.resource_limits['W1']
        print(f"  W1 Limits: {limit}")
