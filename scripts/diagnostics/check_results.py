import pandas as pd

def check_results():
    try:
        df_inv = pd.read_excel('Results/Master_Plan.xlsx', sheet_name='Investments')
        print("Investments up to 2027:")
        print(df_inv[df_inv['Year'] <= 2027])
    except Exception as e:
        print("Error reading excel:", e)

if __name__ == "__main__":
    check_results()
