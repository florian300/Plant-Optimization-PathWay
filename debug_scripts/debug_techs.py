import pprint
from ingestion import PathFinderParser

data = PathFinderParser('PathFinder input.xlsx').parse()

print("--- TECHNOLOGIES ---")
for t_id, tech in data.technologies.items():
    print(f"Tech: {t_id}")
    print(f"  CAPEX: {tech.capex}")
    print(f"  OPEX:  {tech.opex}")
    print(f"  Time:  {tech.implementation_time}")
    print(f"  Impacts:")
    for res_id, imp in tech.impacts.items():
         print(f"    {res_id}: {imp}")
