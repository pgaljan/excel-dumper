# Dumper

A Python script that extracts all non-null rows from Excel worksheets and exports them to CSV format with worksheet names included.

## Features

- **Auto-discovery**: Automatically finds and processes the newest Excel file in the current directory
- **Multi-sheet support**: Extracts data from all worksheets in an Excel file
- **Hidden sheet control**: Option to include or skip hidden worksheets
- **Smart filtering**: Only exports rows containing non-null data
- **Worksheet identification**: Prepends worksheet name as the first column in output
- **Collision-safe naming**: Automatically adds incremental numbers if output files already exist
- **Custom output directory**: Specify where to save the exported CSV files
- **Multiple Excel formats**: Supports .xlsx, .xls, .xlsm, and .xlsb files
- **Timestamped output**: Uses file modification time for consistent naming (ISO 8601 format)

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

### Specify Output Directory

Save the CSV file to a specific directory:

```bash
python dumper.py -output "./exports"
python dumper.py -output "C:\Reports"
```

### Combined Options

```bash
python dumper.py -file "report.xlsx" -output "./exports" -no-hide
```

### Get Help

Display detailed help information:

```bash
python dumper.py -help
```

## Command Line Options

| Option | Description | Example |
|--------|-------------|---------|
| `-file FILE` | Specify Excel file to process | `-file "data.xlsx"` |
| `-no-hide` | Skip hidden worksheets | `-no-hide` |
| `-output DIR` | Output directory for CSV file | `-output "./exports"` |
| `-help` | Show help message | `-help` |

## Output

### File Naming Convention

Output files are named using the pattern:
```
dumper_[original_filename]_[timestamp].csv
```

- **Timestamp format**: ISO 8601 with colons replaced by hyphens for filename compatibility
- **Example**: `2025-07-21T14-30-52-05-00` (July 21, 2025 at 2:30:52 PM, GMT-5)
- **Collision handling**: If file exists, adds `(1)`, `(2)`, etc.

### Example Output Filenames

```
dumper_Sales_Report_2025-07-21T14-30-52-05-00.csv
dumper_Sales_Report_2025-07-21T14-30-52-05-00(1).csv
dumper_Inventory_Data_2025-07-20T09-15-30-05-00.csv
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
- **Directory creation**: Automatically creates output directories if they don't exist

## Dependencies

### Required Python Packages

- **pandas** - Data manipulation and Excel reading
- **openpyxl** - Excel 2007+ (.xlsx) file support and hidden sheet detection
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
Data successfully exported to: dumper_Q3_Report_2025-07-21T14-30-52-05-00.csv
Total rows exported: 1247
```

### Process Specific File with Custom Output

```bash
C:\Data> python dumper.py -file "Annual_Summary.xlsx" -output "./reports"
Extracting data from: Annual_Summary.xlsx
Including hidden sheets: True
Data successfully exported to: ./reports/dumper_Annual_Summary_2025-07-21T14-31-15-05-00.csv
Total rows exported: 892
```

### Skip Hidden Sheets

```bash
C:\Data> python dumper.py -no-hide
Processing newest Excel file: Complex_Workbook.xlsx
Extracting data from: Complex_Workbook.xlsx
Including hidden sheets: False
Skipping hidden sheet: CalculationSheet
Data successfully exported to: dumper_Complex_Workbook_2025-07-21T14-32-01-05-00.csv
Total rows exported: 564
```

### File Collision Handling

```bash
C:\Data> python dumper.py -file "report.xlsx"
# First run creates: dumper_report_2025-07-21T14-30-52-05-00.csv

C:\Data> python dumper.py -file "report.xlsx"
# Second run creates: dumper_report_2025-07-21T14-30-52-05-00(1).csv

C:\Data> python dumper.py -file "report.xlsx"
# Third run creates: dumper_report_2025-07-21T14-30-52-05-00(2).csv
```

## Cross-Platform Compatibility

This script works on Windows, macOS, and Linux:

- **Path handling**: Uses `pathlib` for cross-platform path compatibility
- **Directory creation**: Automatically handles different path separators
- **Timezone handling**: Uses system timezone information
- **File operations**: Compatible across all operating systems

### Platform-Specific Examples

**Windows:**
```cmd
python dumper.py -file "C:\Reports\data.xlsx" -output "C:\Exports"
```

**macOS/Linux:**
```bash
python dumper.py -file "/home/user/data.xlsx" -output "/home/user/exports"
```

## Troubleshooting

### Common Issues

**"Error: Required libraries not found"**
- Install required packages: `pip install pandas openpyxl xlrd`
- Verify installation: `python -c "import pandas, openpyxl; print('Dependencies OK')"`

**"No Excel files found in current directory"**
- Ensure you're in the correct directory
- Or specify a file with `-file` option
- Check for supported file extensions (.xlsx, .xls, .xlsm, .xlsb)

**"Permission denied" or file writing errors**
- Check that the output directory is writable
- Close any Excel files that might be locking the directory
- Try running with administrator privileges if necessary

**"Error reading Excel file"**
- Ensure the Excel file isn't corrupted
- Try opening the file in Excel first to verify it's accessible
- Check that the file isn't password-protected

**Hidden sheet detection not working**
- Hidden sheet detection only works for .xlsx and .xlsm files
- For .xls files, all sheets are processed regardless of `-no-hide` option

### Getting Help

For additional help or to report issues:

1. Run `python dumper.py -help` for detailed usage information
2. Check that all dependencies are properly installed
3. Verify Python version compatibility (3.6+)
4. Ensure you have read permissions for the Excel file and write permissions for the output directory

## Advanced Usage

### Batch Processing

Process multiple files by running the script multiple times:

```bash
for file in *.xlsx; do python dumper.py -file "$file" -output "./processed"; done
```

### Automation

The script can be integrated into automated workflows:

```bash
# Daily report processing
python dumper.py -output "./daily_exports/$(date +%Y-%m-%d)"
```

### Large Files

For very large Excel files:
- The script processes sheets one at a time to manage memory usage
- Progress is shown for each sheet being processed
- Consider using the `-no-hide` option to skip unnecessary hidden sheets