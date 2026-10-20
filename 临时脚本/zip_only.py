import zipfile, glob, os
out_dir = '/Users/a171325./Documents/landrop/运费_处理后/'
zip_path = '/Users/a171325./Documents/landrop/运费_处理后_win兼容.zip'
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    for f in glob.glob(out_dir + '*.xlsx'):
        zf.write(f, arcname=os.path.basename(f))
print(f"Zip created at {zip_path}")
