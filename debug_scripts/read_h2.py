from ingestion import PathFinderParser
import sys

parser = PathFinderParser('PathFinder input.xlsx')
try:
    data = parser.parse()
    print("Entities:", list(data.entities.keys()))
    for e_id, ent in data.entities.items():
        print(f"Base Consumptions: {ent.base_consumptions}")
        print(f"Base Emissions: {ent.base_emissions}")
        
        for p_id, p in ent.processes.items():
            print(f"Process {p_id}:")
            print(f"  Consumption shares: {p.consumption_shares}")
            print(f"  Valid techs: {p.valid_technologies}")
except Exception as e:
    print("Error:", e)
