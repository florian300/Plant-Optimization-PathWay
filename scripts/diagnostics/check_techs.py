import os
import sys

repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
sys.path.append(os.path.join(repo_root, 'src'))

from pathway.core.ingestion import PathFinderParser

def main():
    parser = PathFinderParser(os.path.join(repo_root, 'data', 'raw', 'excel', 'PathFinder input.xlsx'))
    data = parser.parse()
    
    techs_to_check = ['ERH', 'PEM_H2', 'FUEL_TO_H2', 'CCS_1', 'CCS', 'CCU']
    print("=== TECH COSTS ===")
    for t_id in techs_to_check:
        tech = data.technologies.get(t_id)
        if tech:
            print(f"{t_id} -> CAPEX: {tech.capex:,.0f} {tech.capex_unit}, OPEX: {tech.opex:,.0f} {tech.opex_unit}")
            for k, v in tech.impacts.items():
                print(f"   Impact {k}: type={v.get('type')}, value={v.get('value')}")
        else:
            print(f"{t_id} -> NOT FOUND")
            
    print("\n=== RESOURCE PRICES (2030) ===")
    for res in ['EN_ELEC', 'EN_FUEL', 'EN_NAT_GAS', 'EN_GREY_H2_C', 'EN_GREEN_H2_C', 'EN_BLUE_H2_C']:
        if res in data.time_series.resource_prices:
            print(f"{res}: {data.time_series.resource_prices[res].get(2030, 'N/A')} EUR")
            
    print("\n=== CARBON PARAMETERS ===")
    print(f"Price (2030): {data.time_series.carbon_prices.get(2030, 'N/A')} EUR/t")
    print(f"Quota PI (2030): {data.time_series.carbon_quotas_pi.get(2030, 'N/A')}")
    print(f"Penalty (2030): {data.time_series.carbon_penalties.get(2030, 'N/A')}")

if __name__ == "__main__":
    main()
