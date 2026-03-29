from ingestion import PathFinderParser
import pulp
parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()
entity = list(data.entities.values())[0]

for t in [2025, 2030, 2050]:
    ca = 0
    for res_id in entity.sold_resources:
        price = data.time_series.resource_prices.get(res_id, {}).get(t, 0)
        qty = abs(entity.base_consumptions.get(res_id, 0))
        ca += price * qty
    print(f"Year {t} - Potential CA: {ca:,.0f} €")
    print(f"Year {t} - Budget (10%): {ca*0.1:,.0f} €")

# Check technology costs
for t_id, tech in data.technologies.items():
    print(f"Tech {t_id} Capex: {tech.capex:,.0f}")
