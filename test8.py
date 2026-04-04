import pandas as pd, os
with pd.ExcelWriter('test.xlsx', engine='openpyxl', mode='w') as w:
    for i in range(2):
        pd.DataFrame({'a':[i]}).to_excel(w, sheet_name='Charts', startrow=i*5)
