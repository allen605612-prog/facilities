import pandas as pd
df = pd.read_excel(r'C:/Users/user/Downloads/113學年度surfacego名稱與機碼.xlsx', sheet_name=0)
print('欄位：', df.columns.tolist())
print('筆數：', len(df))
print(df.head(3).to_string())
