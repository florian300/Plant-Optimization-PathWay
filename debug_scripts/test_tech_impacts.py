from core.ingestion import PathFinderParser

def debug_tech():
    parser = PathFinderParser('PathFinder input.xlsx')
    data = parser.parse()
    
    print("\n--- Resources ---")
    for r_id, res in data.resources.items():
        if 'ELEC' in r_id or 'FUEL' in r_id:
            print(f"[{r_id}] Unit: {res.unit}")

if __name__ == "__main__":
    debug_tech()
