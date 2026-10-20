import glob
import os
import zipfile
import subprocess

in_dir = '/Users/a171325./Documents/landrop/运费_extracted/运费/'
out_dir = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/'
script = '/Users/a171325./Documents/爬虫/收支流水自动化工具/income-expense-processor/scripts/process_excel.py'

# Process files
for f in glob.glob(in_dir + '*.xlsx'):
    base = os.path.splitext(os.path.basename(f))[0]
    out_file = os.path.join(out_dir, f"{base}_已处理.xlsx")
    print(f"Processing {base}...")
    subprocess.run(['python3', script, f, out_file], check=True)

print("Processing complete. Now zipping...")

# Create Zip compatible with Windows
zip_path = '/Users/a171325./Documents/爬虫/收支流水自动化工具/处理结果归档/运费_处理后_win兼容.zip'
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in glob.glob(out_dir + '*.xlsx'):
        # For GBK compatibility on older Windows, we can manually encode the filename, 
        # but modern Python sets the utf-8 flag which is perfect for Win10+.
        # We will just write it normally, Python handles the UTF-8 bit flag automatically.
        zf.write(f, arcname=os.path.basename(f))

print(f"Zip created at {zip_path}")
