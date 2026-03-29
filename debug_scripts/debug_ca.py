import pulp
from ingestion import PathFinderParser
from optimizer import PathFinderOptimizer

parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()
# Ensure budget is 0 to let it solve and see the baseline revenue
list(data.entities.values())[0].ca_percentage_limit = 0.0

opt = PathFinderOptimizer(data)
opt.build_model()
opt.solve()

print("Budget limit parameter:", list(data.entities.values())[0].ca_percentage_limit)
for res_id in list(data.entities.values())[0].sold_resources:
    v = opt.cons_vars[(2025, res_id)].varValue
    p = data.time_series.resource_prices.get(res_id, {}).get(2025, 0)
    print(f"Resource {res_id}: Cons {v}, Price {p} -> Rev {-v*p:,.0f} €")
