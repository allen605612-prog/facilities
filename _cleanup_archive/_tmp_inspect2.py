import pandas as pd
import sys

df = pd.read_excel(r'C:/Users/user/Downloads/113學年度surfacego名稱與機碼.xlsx', sheet_name=0)
cols = df.columns.tolist()
with open('_tmp_output.txt', 'w', encoding='utf-8') as f:
    f.write(f"欄位: {cols}\n")
    f.write(f"筆數: {len(df)}\n")
    f.write(df.head(5).to_string())
    f.write("\n")
