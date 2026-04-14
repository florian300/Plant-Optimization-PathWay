from pathlib import Path

root = Path(r"c:\Users\flori\Documents\TFM\PathWay_Python_Tool - v2\artifacts\reports\BUSINESS AS USUAL\charts")
files = [
    "Carbon_Prices.json",
    "Carbon_Tax.json",
    "CO2_Trajectory.json",
    "Energy_Mix.json",
    "Financing.json",
    "Indirect_Emissions.json",
    "Investment_Plan.json",
    "Resources_Opex.json",
    "Data_Used.json",
    "Transition_Cost.json",
    "Total_Annual_Opex.json",
    "CO2_Abatement.json"
]

for f in files:
    p = root / f
    print(f"{f}: {p.exists()}")
