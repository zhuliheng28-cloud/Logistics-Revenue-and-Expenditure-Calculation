import pandas as pd
import glob
import os
import zipfile

in_dir = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费_重新拆分后/'
out_dir = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费_已填入模板/'
os.makedirs(out_dir, exist_ok=True)

# Process files
for f in glob.glob(in_dir + '*.xlsx'):
    base = os.path.basename(f)
    print(f"Cleaning {base}...")
    
    # Read Sheet2 from the processed files
    df_src = pd.read_excel(f, sheet_name='Sheet2', dtype=str)
    
    # Create a new DataFrame with template columns
    df_out = pd.DataFrame()
    df_out['运单号'] = ''
    df_out['订单号'] = df_src['订单号']
    
    if '渠道流水单号' in df_src.columns:
        df_out['交易号'] = df_src['渠道流水单号']
    else:
        df_out['交易号'] = ''
        
    if '变动金额' in df_src.columns:
        # Convert to float and take absolute value for "真实运费"
        df_out['真实运费(￥)'] = pd.to_numeric(df_src['变动金额'], errors='coerce').abs()
    else:
        df_out['真实运费(￥)'] = ''
        
    if '币种' in df_src.columns:
        df_out['运费币种（CNY、USD、EUR等）'] = df_src['币种']
    else:
        df_out['运费币种（CNY、USD、EUR等）'] = 'CNY'
        
    df_out['真实重量'] = ''
    df_out['重量单位(kg/g)'] = ''
    
    out_file = os.path.join(out_dir, base)
    df_out.to_excel(out_file, index=False)

print("Cleaning complete. Now zipping...")

zip_path = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/订单拆分后文件.zip'
if os.path.exists(zip_path):
    os.remove(zip_path) # Remove the old one if it exists

with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in glob.glob(out_dir + '*.xlsx'):
        zf.write(f, arcname=os.path.basename(f))

print(f"Zip created at {zip_path}")
