import json
import re
import base64
import numpy as np

html_path = r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\artifacts\reports\Results\results_dashboard.html"

def decode_plotly_data(trace_data):
    if isinstance(trace_data, list):
        return trace_data
    if isinstance(trace_data, dict) and 'bdata' in trace_data:
        dtype = trace_data.get('dtype', 'f8')
        bdata = trace_data.get('bdata', '')
        try:
            return np.frombuffer(base64.b64decode(bdata), dtype=dtype).tolist()
        except:
            return []
    return []

with open(html_path, 'r', encoding='utf-8') as f:
    content = f.read()

match = re.search(r'const dashboardData = (\{.*?\});', content, re.DOTALL)
if match:
    data = json.loads(match.group(1))
    
    if 'entities' in data and 'TOTALENERGIES' in data['entities']:
        sc = data['entities']['TOTALENERGIES']['scenarios']
        for sn in ['BUSINESS AS USUAL', 'CLEAN TECH', 'LOW CARBON BREAK']:
            sd = sc.get(sn, {})
            print(f"\nScenario: {sn}")
            
            # Transition Cost
            if 'graphs' in sd and 'transition_cost' in sd['graphs']:
                fig = sd['graphs']['transition_cost'].get('figure', {})
                print(f"    Transition Cost Highlights:")
                for trace in fig.get('data', []):
                    name = trace.get('name', 'unnamed')
                    y = decode_plotly_data(trace.get('y', []))
                    if y:
                        y_clean = [v for v in y if v is not None and not np.isnan(v)]
                        if y_clean:
                            if "Net Transition Balance" in name:
                                print(f"      {name}: {y_clean[-1]:.2f} M€ (Last)")
                            elif any(k in name for k in ["Carbon Tax", "CAPEX", "OPEX", "Resource"]):
                                print(f"      {name}: {sum(y_clean):.2f} M€ (Total)")

            # Carbon Tax breakdown
            if 'graphs' in sd and 'carbon_tax' in sd['graphs']:
                fig = sd['graphs']['carbon_tax'].get('figure', {})
                print(f"    Carbon Tax Breakdown:")
                for trace in fig.get('data', []):
                    name = trace.get('name', 'unnamed')
                    y = decode_plotly_data(trace.get('y', []))
                    if y:
                        y_clean = [v for v in y if v is not None and not np.isnan(v)]
                        if y_clean:
                             print(f"      {name}: {sum(y_clean):.2f} M€")

            # Investment Limits
            if 'graphs' in sd and 'investment_plan' in sd['graphs']:
                fig = sd['graphs']['investment_plan'].get('figure', {})
                print(f"    Investment Limits:")
                for trace in fig.get('data', []):
                    name = trace.get('name', 'unnamed')
                    if "Limit" in name:
                        y = decode_plotly_data(trace.get('y', []))
                        if y:
                            y_clean = [v for v in y if v is not None and not np.isnan(v)]
                            if y_clean:
                                 print(f"      {name}: {y_clean[-1]:.2f} M€ (2050)")
