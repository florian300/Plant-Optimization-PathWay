from ingestion import PathFinderParser
parser = PathFinderParser('PathFinder input.xlsx')
data = parser.parse()
entity = list(data.entities.values())[0]
print("Parsed CA % Limit:", entity.ca_percentage_limit)
