import json
import base64
import struct

def decode_plotly_bdata(bdata, dtype):
    raw = base64.b64decode(bdata)
    if dtype == 'f8':
        fmt = f'<{len(raw)//8}d'
        return struct.unpack(fmt, raw)
    return []

with open('c:/Users/flori/Documents/TFM/PathWay_Python_Tool - v2/artifacts/reports/TOTALENERGIES/CLEAN TECH/charts/Transition_Cost.json', 'r') as f:
    d = json.load(f)

for trace in d['data']:
    name = trace.get('name', 'Unknown')
    y = trace.get('y', [])
    if isinstance(y, dict) and 'bdata' in y:
        y_vals = decode_plotly_bdata(y['bdata'], y['dtype'])
    else:
        y_vals = y
    
    if y_vals:
        print(f"Trace: {name}")
        print(f"  2028: {y_vals[3]:.2f}")
        print(f"  2029: {y_vals[4]:.2f}")
        print(f"  Diff: {y_vals[4] - y_vals[3]:.2f}")
