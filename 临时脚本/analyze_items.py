import pandas as pd
import glob
import os

res = set()
for f in glob.glob('/Users/a171325./Documents/landrop/运费_extracted/运费/*.xlsx'):
    df = pd.read_excel(f, dtype=str)
    if '费用项' in df.columns:
        for _, row in df[['收支类型', '费用项', '变动金额']].drop_duplicates().iterrows():
            typ = str(row['收支类型'])
            item = str(row['费用项'])
            amt = str(row['变动金额'])
            # Check sign
            if amt.startswith('-'):
                sign = '-'
            elif amt == '0' or amt == '0.0' or amt == 'nan':
                sign = '0'
            else:
                sign = '+'
            res.add((typ, item, sign))

print("收支类型 | 费用项 | 金额正负")
print("-" * 50)
for x in sorted(list(res)):
    print(f"{x[0]} | {x[1]} | {x[2]}")
