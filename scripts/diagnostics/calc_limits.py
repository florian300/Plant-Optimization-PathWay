from ingestion import PathFinderParser
import sys

parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()

entity = list(data.entities.values())[0]
print(f"\n--- BASELINE FOR {entity.id} ---")
print(f"Base Emissions: {entity.base_emissions:,.2f} tCO2")
print(f"Base EN_FUEL Consumption: {entity.base_consumptions.get('EN_FUEL', 0):,.2f} MW")

t = 2026
print(f"\n--- CALCULATION FOR YEAR {t} ---")

ca_yr = 0.0
for res_id in entity.sold_resources:
    price = data.time_series.resource_prices.get(res_id, {}).get(t, 0.0)
    # production is assumed positive or negative, check ingestion
    # actually base_consumptions for sold items are usually negative in this model
    cons = entity.base_consumptions.get(res_id, 0.0)
    ca_yr += price * (-cons) if cons < 0 else price * cons 
    
budget_yr = ca_yr * entity.ca_percentage_limit
print(f"Estimated CA limit factor: {entity.ca_percentage_limit}")
print(f"Estimated CA for {t}: {ca_yr:,.2f} Euros")
print(f"1-Year Budget Limit (CA%) for {t}: {budget_yr:,.2f} Euros")
print(f"5-Year Rolling Budget Max Capacity: {budget_yr * 5:,.2f} Euros")

print("\n--- TECHNOLOGY COSTS IF APPLIED IN 2026 ---")
for tech_id, tech in data.technologies.items():
    cap_capex = 1.0
    if tech.capex_per_unit:
        if tech.capex_unit == 'tCO2': cap_capex = entity.base_emissions
        elif 'MW' in tech.capex_unit.upper(): cap_capex = entity.base_consumptions.get('EN_FUEL', 0.0)
        
    cap_opex = 1.0
    if tech.opex_per_unit:
        if tech.opex_unit == 'tCO2': cap_opex = entity.base_emissions
        elif 'MW' in tech.opex_unit.upper(): cap_opex = entity.base_consumptions.get('EN_FUEL', 0.0)
        
    total_capex = tech.capex * cap_capex
    total_opex = tech.opex * cap_opex
    
    print(f"\nTechnology: {tech_id}")
    unit_str = tech.capex_unit if tech.capex_per_unit else 'Fixed'
    print(f"  Multiplier Size: {cap_capex:,.2f} {unit_str}")
    print(f"  Base CAPEX rate: {tech.capex:,.2f} / Base OPEX rate: {tech.opex:,.2f}")
    print(f"  -> TOTAL CAPEX COST: {total_capex:,.2f} Euros")
    print(f"  -> TOTAL OPEX COST / yr: {total_opex:,.2f} Euros")
    print(f"  Feasible under 1-yr Budget of {budget_yr:,.2f}?: {'YES' if total_capex <= budget_yr else 'NO'}")
    print(f"  Feasible under 5-yr Max Budget of {budget_yr * 5:,.2f}?: {'YES' if total_capex <= budget_yr * 5 else 'NO'}")
