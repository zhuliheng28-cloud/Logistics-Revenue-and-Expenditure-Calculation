import pandas as pd
import glob
import os
import subprocess
import zipfile
import shutil

in_dir = '/Users/a171325./Documents/landrop/运费7月_extracted/运费7月/'
process_dir = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费7月_重新拆分后/'
template_dir = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费7月_已填入模板/'
script = '/Users/a171325./Documents/爬虫/收支流水自动化工具/income-expense-processor/scripts/process_excel.py'
zip_path = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费7月订单拆分后文件.zip'

os.makedirs(process_dir, exist_ok=True)
os.makedirs(template_dir, exist_ok=True)

print("Step 1: Process and Split")
for f in glob.glob(in_dir + '*.xlsx'):
    base = os.path.basename(f)
    out_file = os.path.join(process_dir, base.replace('.xlsx', '_已处理.xlsx'))
    print(f"  -> Processing {base}...")
    subprocess.run(['python3', script, f, out_file], check=True)

print("\nStep 2: Clean into Template")
for f in glob.glob(process_dir + '*.xlsx'):
    base = os.path.basename(f)
    print(f"  -> Cleaning {base}...")
    df_src = pd.read_excel(f, sheet_name='Sheet2', dtype=str)
    
    df_out = pd.DataFrame()
    df_out['运单号'] = ''
    df_out['订单号'] = df_src['订单号']
    
    if '渠道流水单号' in df_src.columns:
        df_out['交易号'] = df_src['渠道流水单号']
    else:
        df_out['交易号'] = ''
        
    if '变动金额' in df_src.columns:
        df_out['真实运费(￥)'] = pd.to_numeric(df_src['变动金额'], errors='coerce').abs()
    else:
        df_out['真实运费(￥)'] = ''
        
    if '币种' in df_src.columns:
        df_out['运费币种（CNY、USD、EUR等）'] = df_src['币种']
    else:
        df_out['运费币种（CNY、USD、EUR等）'] = 'CNY'
        
    df_out['真实重量'] = ''
    df_out['重量单位(kg/g)'] = ''
    
    out_file = os.path.join(template_dir, base)
    df_out.to_excel(out_file, index=False)

print("\nStep 3: Zipping")
if os.path.exists(zip_path):
    os.remove(zip_path)

with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in glob.glob(template_dir + '*.xlsx'):
        zf.write(f, arcname=os.path.basename(f))

print(f"\nAll done! Zip created at {zip_path}")
