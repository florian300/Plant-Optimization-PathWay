import pandas as pd
import sys
try:
    df_invest = pd.read_excel('Results/Master_Plan.xlsx', sheet_name='Investments')
    print("--- INVESTMENTS ---")
    print(df_invest)
    
    df_mix = pd.read_excel('Results/Master_Plan.xlsx', sheet_name='Energy_Mix')
    print("\n--- ENERGY MIX (Ends of simulation) ---")
    print(df_mix.tail(5))
    
    print("\n--- SUM OF H2 MIX ---")
    h2_cols = [c for c in df_mix.columns if 'H2' in c.upper()]
    for c in h2_cols:
        print(f"{c}: {df_mix[c].sum()}")
        
    df_indir = pd.read_excel('Results/Master_Plan.xlsx', sheet_name='Indirect_Emissions')
    print("\n--- INDIRECT EMISSIONS (Ends of simulation) ---")
    print(df_indir.tail(5))
except Exception as e:
    print("Error:", e)
    sys.exit(1)
