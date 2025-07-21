# Dumper 

Simple python workflow leveraging [sheet-excavator](https://pypi.org/project/sheet-excavator/) to extract data from Excel worksheets to CSV

**USAGE**:
    `python dumper.py [OPTIONS]`

**OPTIONS**:

`-file` FILE          Specify Excel file to process (default: newest Excel file in current directory)

`-no-hide`           Skip hidden worksheets (default: include all worksheets)

`-help`              Show this help message

**EXAMPLES**:

`python dumper.py`                    # Process newest Excel file, include all sheets

`python dumper.py -file data.xlsx`    # Process specific file

`python dumper.py -no-hide`           # Skip hidden worksheets

`python dumper.py -file data.xlsx -no-hide`  # Specific file, skip hidden sheets

**OUTPUT**:
Creates a CSV file named "dumper_[original_filename]_[timestamp].csv" with:
- Timestamp is the last modified time of the originating Excel file
- Timestamp format: YYYYMMDD_HHMMSS_TZ (e.g., dumper_data_20250721_143052_EST.csv)
- First column: Worksheet name
- Remaining columns: Original data from worksheets
- Only non-empty rows are included

**PYTHON DEPENDENCIES**:
- sheet-excavator  (pip install sheet-excavator)
- Standard library: argparse, csv, glob, os, sys, pathlib, datetime

**SUPPORTED EXCEL FORMATS**:
- .xlsx (Excel 2007+)
- .xls  (Excel 97-2003)
- .xlsm (Excel Macro-Enabled)
- .xlsb (Excel Binary)
