from ingestion import PathFinderParser

parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()

for r_id, factors in data.time_series.other_emissions_factors.items():
    if 'H2' in r_id.upper():
        print(f"--- {r_id} ---")
        print(factors.get(2025))
        print(factors.get(2050))
