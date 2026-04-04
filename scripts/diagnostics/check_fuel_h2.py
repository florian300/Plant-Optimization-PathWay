from ingestion import PathFinderParser
parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()
tech = data.technologies['FUEL_TO_H2']

print("FUEL_TO_H2 Capex:", tech.capex)
print("FUEL_TO_H2 Capex unit:", tech.capex_unit)
print("FUEL_TO_H2 Capex per unit:", tech.capex_per_unit)
