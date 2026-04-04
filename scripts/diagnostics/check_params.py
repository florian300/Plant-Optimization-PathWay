import pandas as pd
try:
    xl = pd.ExcelFile('PathFinder input.xlsx')
    df_overview = xl.parse('OverView', header=None)
    for i, row in df_overview.iterrows():
        row_vals = [str(x).strip().upper() if pd.notna(x) else "" for x in row]
        if "DURATION SIMULATION (S)" in row_vals or "DURURATION SIMULATION (S)" in row_vals:
            keyword = "DURURATION SIMULATION (S)" if "DURURATION SIMULATION (S)" in row_vals else "DURATION SIMULATION (S)"
            idx = row_vals.index(keyword)
            print(f"{keyword}: {row.iloc[idx + 1]}")
        if "ERROR SIMULATION (%)" in row_vals:
            idx = row_vals.index("ERROR SIMULATION (%)")
            print(f"ERROR SIMULATION (%): {row.iloc[idx + 1]}")
except Exception as e:
    print(f"Error: {e}")
