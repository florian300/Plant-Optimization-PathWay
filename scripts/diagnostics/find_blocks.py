import pandas as pd
import json

file_path = 'PathFinder input.xlsx'
xl = pd.ExcelFile(file_path)
blocks = {}

for sheet_name in xl.sheet_names:
    df = xl.parse(sheet_name, header=None)
    sheet_blocks = []
    for i, row in df.iterrows():
        for j, val in enumerate(row):
            if isinstance(val, str) and val.strip().upper() in ['START', 'END']:
                # The block name is usually in the preceding column or same column
                # Let's just grab the whole row up to the 'START'/'END'
                prefix = [str(x) for x in row[:j] if pd.notna(x) and str(x).strip() != '']
                sheet_blocks.append({
                    'row': i,
                    'col': j,
                    'type': val.strip().upper(),
                    'prefix': prefix
                })
    blocks[sheet_name] = sheet_blocks

with open('blocks_structure.json', 'w', encoding='utf-8') as f:
    json.dump(blocks, f, indent=2, ensure_ascii=False)
print("Dumped blocks to blocks_structure.json")
