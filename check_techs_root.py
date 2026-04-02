import os
import sys

base_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(base_path, 'src'))

from core.ingestion import PathFinderParser

def main():
    parser = PathFinderParser(os.path.join(base_path, 'PathFinder input.xlsx'))
    data = parser.parse()
    
    techs_to_check = ['ERH', 'PEM_H2', 'FUEL_TO_H2', 'CCS_1', 'CCS', 'CCU']
    print("=== TECH COSTS ===")
    for t_id in techs_to_check:
        tech = data.technologies.get(t_id)
        if tech:
            print(f"{t_id} -> CAPEX_base: {tech.capex:,.0f} {tech.capex_unit}, OPEX_base: {tech.opex:,.0f} {tech.opex_unit}")
            print(f"   2030: CAPEX={tech.capex_by_year.get(2030, tech.capex):,.0f}, OPEX={tech.opex_by_year.get(2030, tech.opex):,.0f}")
            print(f"   2050: CAPEX={tech.capex_by_year.get(2050, tech.capex):,.0f}, OPEX={tech.opex_by_year.get(2050, tech.opex):,.0f}")
            for k, v in tech.impacts.items():
                print(f"   Impact {k}: type={v.get('type')}, value={v.get('value')} (ref: {v.get('ref_resource')})")
        else:
            print(f"{t_id} -> NOT FOUND")
            
    print("\n=== RESOURCE PRICES & EMISSIONS (2025 to 2050) ===")
    for res in data.resources:
        if 'H2' in res.upper() and res in data.time_series.resource_prices:
            p2030 = data.time_series.resource_prices[res].get(2030, 'N/A')
            p2050 = data.time_series.resource_prices[res].get(2050, 'N/A')
            f2030 = data.time_series.other_emissions_factors.get(res, {}).get(2030, 0.0)
            f2050 = data.time_series.other_emissions_factors.get(res, {}).get(2050, 0.0)
            print(f"{res}: 2030 Price={p2030} EUR, factor={f2030} t/MWh")
            print(f"         2050 Price={p2050} EUR, factor={f2050} t/MWh")
            
    print("\n=== SYSTEM & GRANT PARAMETERS ===")
    if data.entities:
        ent = next(iter(data.entities.values()))
        print(f"Entity {ent.id} Operating Hours: {ent.annual_operating_hours} h/year")
    print(f"Grant Rate: {data.grant_params.rate*100}%")
    print(f"Grant Cap: {data.grant_params.cap:,.0f} EUR")
    
    print("\n=== CARBON PARAMETERS ===")
    print(f"Price (2030): {data.time_series.carbon_prices.get(2030, 'N/A')} EUR/t")
    print(f"Price (2050): {data.time_series.carbon_prices.get(2050, 'N/A')} EUR/t")
    print(f"Penalty (2030): {data.time_series.carbon_penalties.get(2030, 'N/A')}")
    print(f"Penalty (2050): {data.time_series.carbon_penalties.get(2050, 'N/A')}")

if __name__ == "__main__":
    main()
