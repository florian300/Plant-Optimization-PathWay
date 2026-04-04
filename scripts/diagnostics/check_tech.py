from ingestion import PathFinderParser
parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()
ccs = data.technologies['CCS']

print("CCS ID:", ccs.id)
print("CCS Capex:", ccs.capex)
print("CCS Capex per unit:", ccs.capex_per_unit)
print("CCS Capex unit:", ccs.capex_unit)
print("CCS Opex:", ccs.opex)
print("CCS Opex per unit:", ccs.opex_per_unit)
print("CCS Opex unit:", ccs.opex_unit)
