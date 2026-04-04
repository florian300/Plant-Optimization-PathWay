from ingestion import PathFinderParser
p = PathFinderParser('PathFinder input.xlsx')
d = p.parse()
ent = list(d.entities.values())[0]
print(f"Entity: {ent.id}")
print(f"Production: {ent.production_level}")
print(f"Base Elec: {ent.base_consumptions.get('EN_ELEC')} MWh")
print(f"Base Fuel: {ent.base_consumptions.get('EN_FUEL')} GJ or MWh")
