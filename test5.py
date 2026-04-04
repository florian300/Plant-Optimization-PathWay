import pandas as pd, os
with pd.ExcelWriter('test.xlsx', engine='openpyxl') as w:
    for i in range(2):
        pd.DataFrame({'a':[i]}).to_excel(w, sheet_name='Sheet1', startrow=i*5)
print(pd.read_excel('test.xlsx'))
