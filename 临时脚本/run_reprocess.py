import glob, os, subprocess, zipfile

in_dir = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费_处理后/'
out_dir = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费_重新拆分后/'
script = '/Users/a171325./Documents/爬虫/收支流水自动化工具/income-expense-processor/scripts/process_excel.py'

os.makedirs(out_dir, exist_ok=True)

# Process files
for f in glob.glob(in_dir + '*.xlsx'):
    base = os.path.splitext(os.path.basename(f))[0].replace('_已处理', '')
    out_file = os.path.join(out_dir, f"{base}_已处理.xlsx")
    print(f"Processing {base}...")
    subprocess.run(['python3', script, f, out_file], check=True)

print("Processing complete. Now zipping...")

zip_path = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/订单拆分后文件.zip'
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in glob.glob(out_dir + '*.xlsx'):
        zf.write(f, arcname=os.path.basename(f))

print(f"Zip created at {zip_path}")
