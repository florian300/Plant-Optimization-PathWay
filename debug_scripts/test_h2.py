from main import *
from ingestion import PathFinderParser
from optimizer import PathFinderOptimizer

parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()

opt = PathFinderOptimizer(data)
opt.build_model()
opt.solve()

t = 2050
print(f"--- 2050 H2 Variables ---")
print("cons_vars['EN_H2_P']:", opt.cons_vars[(t, 'EN_H2_P')].varValue)
for k, v in opt.h2_supply_vars[t].items():
    print(f"supply_vars[{k}]:", v.varValue)

print("h2_demand affine expression eval:", sum([opt.cons_vars[(t, r_id)].varValue for r_id in ['EN_H2_P', 'EN_H2_C'] if r_id in data.resources]))
