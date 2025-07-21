# Dumper

A Python script that extracts all non-null rows from Excel worksheets and exports them to CSV format with worksheet names included.

## Features

- **Auto-discovery**: Automatically finds and processes the newest Excel file in the current directory
- **Multi-sheet support**: Extracts data from all worksheets in an Excel file
- **Smart filtering**: Only exports rows containing non-null data
- **Worksheet identification**: Prepends worksheet name as the first column in output
- **Collision-safe naming**: Automatically adds incremental numbers if output files already exist
- **Multiple Excel formats**: Supports .xlsx, .xls, .xlsm, and .xlsb files
- **Timestamped output**: Uses file modification time for consistent naming

## Installation

### Prerequisites

Python 3.6 or higher is required.

### Install Dependencies

```bash
pip install pandas openpyxl xlrd
```

Or if you're using a specific Python installation:

```bash
python -m pip install pandas openpyxl xlrd
```

### Download the Script

Save the `dumper.py` file to your desired directory.

## Usage

### Basic Usage

Process the newest Excel file in the current directory:

```bash
python dumper.py
```

### Specify a File

Process a specific Excel file:

```bash
python dumper.py -file "data.xlsx"
python dumper.py -file "C:\Path\To\File.xls"
```

### Skip Hidden Worksheets

Exclude hidden worksheets from processing:

```bash
python dumper.py -no-hide
```

### Combined Options

```bash
python dumper.py -file "report.xlsx" -no-hide
```

### Get Help

Display detailed help information:

```bash
python dumper.py -help
```

## Output

### File Naming Convention

Output files are named using the pattern:
```
dumper_[original_filename]_[timestamp].csv
```

- **Timestamp format**: ISO 8601-like, Windows-compatible (e.g., `2025-07-21T14-30-52-0500`)
- **Collision handling**: If file exists, adds `(1)`, `(2)`, etc.

### Example Output Filenames

```
dumper_Sales_Report_2025-07-21T14-30-52-0500.csv
dumper_Sales_Report_2025-07-21T14-30-52-0500(1).csv
dumper_Inventory_Data_2025-07-20T09-15-30-0500.csv
```

### CSV Structure

The output CSV contains:

1. **First column**: Worksheet name
2. **Remaining columns**: Original data from Excel worksheets (`Column_1`, `Column_2`, etc.)
3. **Header row**: `Worksheet, Column_1, Column_2, ...`
4. **Data rows**: Only non-empty rows from the source Excel file

### Sample Output

```csv
Worksheet,Column_1,Column_2,Column_3
Sheet1,John Doe,Sales Manager,50000
Sheet1,Jane Smith,Developer,65000
Summary,Total Employees,,2
Summary,Average Salary,,57500
```

## Supported Excel Formats

- **.xlsx** - Excel 2007+ (Open XML Format)
- **.xls** - Excel 97-2003 (Binary Format)
- **.xlsm** - Excel Macro-Enabled Workbook
- **.xlsb** - Excel Binary Workbook

## Error Handling

The script provides clear error messages for common issues:

- **File not found**: When specified file doesn't exist
- **No Excel files**: When no Excel files are found in directory
- **Import errors**: When required libraries are missing
- **Read errors**: When Excel file is corrupted or inaccessible
- **Write errors**: When output location is not writable

## Dependencies

### Required Python Packages

- **pandas** - Data manipulation and Excel reading
- **openpyxl** - Excel 2007+ (.xlsx) file support
- **xlrd** - Legacy Excel (.xls) file support

### Standard Library Modules

- `argparse` - Command line argument parsing
- `csv` - CSV file writing
- `glob` - File pattern matching
- `os` - Operating system interface
- `sys` - System-specific parameters
- `datetime` - Date and time handling
- `pathlib` - Object-oriented filesystem paths

## Examples

### Process Latest File

```bash
C:\Data> python dumper.py
Processing newest Excel file: Q3_Report.xlsx
Extracting data from: Q3_Report.xlsx
Including hidden sheets: True
Data successfully exported to: dumper_Q3_Report_2025-07-21T14-30-52-0500.csv
Total rows exported: 1247
```

### Process Specific File

```bash
C:\Data> python dumper.py -file "Annual_Summary.xlsx"
Extracting data from: Annual_Summary.xlsx
Including hidden sheets: True
Data successfully exported to: dumper_Annual_Summary_2025-07-21T14-31-15-0500.csv
Total rows exported: 892
```

### Skip Hidden Sheets

```bash
C:\Data> python dumper.py -no-hide
Processing newest Excel file: Complex_Workbook.xlsx
Extracting data from: Complex_Workbook.xlsx
Including hidden sheets: False
Data successfully exported to: dumper_Complex_Workbook_2025-07-21T14-32-01-0500.csv
Total rows exported: 564
```

## Troubleshooting

### Common Issues

**"Error: pandas library not found"**
- Install required packages: `pip install pandas openpyxl xlrd`

**"No Excel files found in current directory"**
- Ensure you're in the correct directory
- Or specify a file with `-file` option

**"Permission denied" or file writing errors**
- Check that the output directory is writable
- Close any Excel files that might be locking the directory

**"Error reading Excel file"**
- Ensure the Excel file isn't corrupted
- Try opening the file in Excel first to verify it's accessible
- Check that the file isn't password-protected

### Getting Help

For additional help or to report issues:

1. Run `python dumper.py -help` for detailed usage information
2. Check that all dependencies are properly installed
3. Verify Python version compatibility (3.6+)
