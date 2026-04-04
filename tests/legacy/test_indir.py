import pandas as pd
try:
    df_indir = pd.read_excel('Results/Master_Plan.xlsx', sheet_name='Indirect_Emissions')
    print("--- INDIRECT EMISSIONS COLUMNS ---")
    print(df_indir.columns.tolist())
    print("\n--- 2025 ---")
    print(df_indir.head(1).to_string())
    print("\n--- 2050 ---")
    print(df_indir.tail(1).to_string())
except Exception as e:
    print(e)
