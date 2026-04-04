import pandas as pd
from ingestion import PathFinderParser

parser = PathFinderParser('PathFinder input.xlsx')
try:
    df = parser.xl.parse('REFINERY', header=None)
    blocks = parser._find_blocks(df)
    for b in blocks:
        if 'REF' in str(b['prefix']).upper() or b['type'] == 'START':
            print(b)
    
    # Let's extract everything inside the REF block if we find it
    ref_start = next((b['row'] for b in blocks if b['type'] == 'START' and 'REF' in str(b['prefix']).upper()), None)
    ref_end = next((b['row'] for b in blocks if b['type'] == 'END' and 'REF' in str(b['prefix']).upper()), None)
    if ref_start is not None and ref_end is not None:
        print(f"\nExtracted REF block from row {ref_start} to {ref_end}:")
        print(df.iloc[ref_start+1:ref_end])
except Exception as e:
    print("Error:", e)
