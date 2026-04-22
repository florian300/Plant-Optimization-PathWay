import pandas as pd

excel_path = 'c:/Users/flori/Documents/TFM/PathWay_Python_Tool - v2/artifacts/reports/TOTALENERGIES/CLEAN TECH/Master_Plan.xlsx'

sheets = ['Technology_Costs', 'Financing', 'Energy_Mix', 'CO2_Trajectory']

for sheet in sheets:
    print(f"\n--- Sheet: {sheet} ---")
    df = pd.read_excel(excel_path, sheet_name=sheet)
    if 'Year' in df.columns:
        row_2028 = df[df['Year'] == 2028]
        row_2029 = df[df['Year'] == 2029]
        print("Year 2028:")
        print(row_2028.to_string())
        print("Year 2029:")
        print(row_2029.to_string())
