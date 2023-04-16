# sms-log-convert
Convert xml output by 'SMS Backup & Restore' to single *.xlsx, including images if available.

## Install

```bash
# Clone the repo
git clone https://github.com/prettydeep/sms-log-convert.git
cd sms-log-convert/

# Install openpyxl
pip install openpyxl
```

## Convert call logs and sms/mms logs from *.xml to *.xlsx

```
python [call/sms]-log-convert.py --input_file /path/to/input.xml --output_file /path/to/output.xlsx
```

## If you have issues with .mpo KeyError
```
pip show openpyxl
```
Modify the [Location]/openpyxl/packaging/manifest.py so that the function below includes the second line:
```
def _register_mimetypes(self, filenames):
        mimetypes.add_type('image/mpo', '.mpo')
        for fn in filenames:
            ext = os.path.splitext(fn)[-1]
            if not ext:
                continue
            mime = mimetypes.types_map[True][ext]
            fe = FileExtension(ext[1:], mime)
            self.Default.append(fe)
```

