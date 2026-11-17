# read_excel
Simple test project to create program and learn.

## Usage

Install the Excel-reading dependency:

```console
python3 -m pip install -r requirements.txt
```

Print every worksheet in an Excel workbook:

```console
python3 read_excel.py path/to/workbook.xlsx
```

Print selected worksheets by name:

```console
python3 read_excel.py path/to/workbook.xlsx --sheet "Sheet1" --sheet "Summary"
```

The program prints tab-separated rows and marks each worksheet with a
`[Sheet: NAME]` heading. It supports `.xlsx` and `.xlsm` files through
`openpyxl`. Use `python3 read_excel.py --help` for command-line help.
