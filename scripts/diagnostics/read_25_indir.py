import pandas as pd
import sys

try:
    df_indir = pd.read_excel('Results/Master_Plan.xlsx', sheet_name='Indirect_Emissions')
    print("--- INDIRECT EMISSIONS ---")
    print(df_indir.head())
except Exception as e:
    print(e)
