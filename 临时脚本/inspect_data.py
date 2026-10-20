import pandas as pd

file_path = '/Users/a171325./Downloads/收支流水20260712230139 (1) 2.xlsx'
df = pd.read_excel(file_path, sheet_name=0)
print("Columns:", list(df.columns))
print("First 2 rows:")
print(df.head(2))
