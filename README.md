# Dumper

Cross-platform Python-based ETL pre-processor that flattens Excel files into a predictable CSV format, preserving key source metadata for easy ingestion into data pipelines.

## Features

- **Auto-discovery**: Automatically finds and processes the newest Excel file in the specified directory
- **Multi-sheet support**: Extracts data from all worksheets in an Excel file
- **Hidden sheet control**: Option to include or skip hidden worksheets
- **Smart filtering**: Only exports rows containing non-null data
- **Worksheet identification**: Prepends worksheet name as the first column in output
- **Collision-safe naming**: Automatically adds incremental numbers if output files already exist
- **Custom output directory**: Specify where to save the exported CSV files
- **Flexible input sources**: Search for files in any directory or process specific files
- **Multiple Excel formats**: Supports .xlsx, .xls, .xlsm, and .xlsb files
- **Timestamped output**: Uses file modification time for consistent naming (ISO 8601 format)
- **Row number tracking**: Option to include Excel row numbers for data traceability

## Installation

### Prerequisites

Python 3.6 or higher is required.

### Install Dependencies

<pre><code>pip install pandas openpyxl xlrd
</code></pre>

Or if you're using a specific Python installation:

<pre><code>bash
python -m pip install pandas openpyxl xlrd
</code></pre>

### Download the Script

Save the `dumper.py` file to your desired directory.

## Usage

### Basic Usage

Process the newest Excel file in the current directory:

<pre><code>python dumper.py
</code></pre>

### Specify Input Directory

Search for and process the newest Excel file in a specific directory:

<pre><code>python dumper.py -input "./data"
python dumper.py -input "C:\Reports"
</code></pre>

### Specify a File

Process a specific Excel file:

<pre><code>python dumper.py -file "data.xlsx"
python dumper.py -file "C:\Path\To\File.xls"
</code></pre>

### Specify File in Input Directory

Process a specific file within an input directory:

<pre><code>python dumper.py -input "./source" -file "report.xlsx"
python dumper.py -input "/data" -file "monthly.xlsx"
</code></pre>

### Include Row Numbers

Include Excel row numbers in the output for data traceability:

<pre><code>python dumper.py -rownumbers
python dumper.py -file "data.xlsx" -rownumbers
</code></pre>

### Skip Hidden Worksheets

Exclude hidden worksheets from processing:

<pre><code>python dumper.py -no-hide
</code></pre>

### Specify Output Directory

Save the CSV file to a specific directory:

<pre><code>python dumper.py -output "./exports"
python dumper.py -output "C:\Reports"
</code></pre>

### Combined Options

<pre><code>python dumper.py -input "./source" -output "./exports" -no-hide
python dumper.py -input "/data" -file "report.xlsx" -output "./processed" -rownumbers
python dumper.py -file "data.xlsx" -output "./exports" -no-hide -rownumbers
</code></pre>

### Get Help

Display detailed help information:

<pre><code>python dumper.py -help
</code></pre>

## Command Line Options

| Option | Description | Example |
|--------|-------------|---------|
| `-file FILE` | Specify Excel file to process | `-file "data.xlsx"` |
| `-input DIR` | Input directory to search for Excel files | `-input "./source"` |
| `-output DIR` | Output directory for CSV file | `-output "./exports"` |
| `-no-hide` | Skip hidden worksheets | `-no-hide` |
| `-rownumbers` | Include Excel row numbers in output | `-rownumbers` |
| `-help` | Show help message | `-help` |

## Output

### File Naming Convention

Output files are named using the pattern:
</code></pre>dumper_[original_filename]_[timestamp].csv
</code></pre>

- **Timestamp format**: ISO 8601 with colons replaced by hyphens for filename compatibility
- **Example**: `2025-07-21T14-30-52-05-00` (July 21, 2025 at 2:30:52 PM, GMT-5)
- **Collision handling**: If file exists, adds `(1)`, `(2)`, etc.

### Example Output Filenames

</code></pre>dumper_Sales_Report_2025-07-21T14-30-52-05-00.csv
dumper_Sales_Report_2025-07-21T14-30-52-05-00(1).csv
dumper_Inventory_Data_2025-07-20T09-15-30-05-00.csv
</code></pre>

### CSV Structure

The output CSV contains:

1. **First column**: Worksheet name
2. **Second column**: Excel row number (if `-rownumbers` option used)
3. **Remaining columns**: Original data from Excel worksheets (`Column_1`, `Column_2`, etc.)
4. **Header row**: `Worksheet, Column_1, Column_2, ...` or `Worksheet, Row_Number, Column_1, Column_2, ...`
5. **Data rows**: Only non-empty rows from the source Excel file

### Sample Output

**Without row numbers:**
</code></pre>csv
Worksheet,Column_1,Column_2,Column_3
Sheet1,John Doe,Sales Manager,50000
Sheet1,Jane Smith,Developer,65000
Summary,Total Employees,,2
Summary,Average Salary,,57500
</code></pre>

**With row numbers (`-rownumbers`):**
</code></pre>csv
Worksheet,Row_Number,Column_1,Column_2,Column_3
Sheet1,2,John Doe,Sales Manager,50000
Sheet1,3,Jane Smith,Developer,65000
Summary,5,Total Employees,,2
Summary,6,Average Salary,,57500
</code></pre>

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

### Process Latest File from Specific Directory

<pre><code>C:\> python dumper.py -input "C:\Data"
Processing newest Excel file: Q3_Report.xlsx
From directory: C:\Data
Extracting data from: C:\Data\Q3_Report.xlsx
Including hidden sheets: True
Including row numbers: False
Data successfully exported to: dumper_Q3_Report_2025-07-21T14-30-52-05-00.csv
Total rows exported: 1247
</code></pre>

### Process Specific File with Input and Output Directories

<pre><code>C:\> python dumper.py -input "./source" -file "Annual_Summary.xlsx" -output "./reports" -rownumbers
Extracting data from: ./source/Annual_Summary.xlsx
Including hidden sheets: True
Including row numbers: True
Data successfully exported to: ./reports/dumper_Annual_Summary_2025-07-21T14-31-15-05-00.csv
Total rows exported: 892
</code></pre>

### Skip Hidden Sheets

<pre><code>C:\Data> python dumper.py -no-hide
Processing newest Excel file: Complex_Workbook.xlsx
Extracting data from: Complex_Workbook.xlsx
Including hidden sheets: False
Skipping hidden sheet: CalculationSheet
Data successfully exported to: dumper_Complex_Workbook_2025-07-21T14-32-01-05-00.csv
Total rows exported: 564
</code></pre>

### File Collision Handling

<pre><code>C:\Data> python dumper.py -file "report.xlsx"
# First run creates: dumper_report_2025-07-21T14-30-52-05-00.csv

C:\Data> python dumper.py -file "report.xlsx"
# Second run creates: dumper_report_2025-07-21T14-30-52-05-00(1).csv

C:\Data> python dumper.py -file "report.xlsx"
# Third run creates: dumper_report_2025-07-21T14-30-52-05-00(2).csv
</code></pre>

## Cross-Platform Compatibility

This script works on Windows, macOS, and Linux:

- **Path handling**: Uses `pathlib` for cross-platform path compatibility
- **Directory creation**: Automatically handles different path separators
- **Timezone handling**: Uses system timezone information
- **File operations**: Compatible across all operating systems

### Platform-Specific Examples

**Windows:**
</code></pre>cmd
python dumper.py -file "C:\Reports\data.xlsx" -output "C:\Exports"
</code></pre>

**macOS/Linux:**
<pre><code>python dumper.py -file "/home/user/data.xlsx" -output "/home/user/exports"
</code></pre>

## Troubleshooting

### Common Issues

**"Error: Required libraries not found"**
- Install required packages: `pip install pandas openpyxl xlrd`
- Verify installation: `python -c "import pandas, openpyxl; print('Dependencies OK')"`

**"No Excel files found in directory"**
- Ensure you're in the correct directory or specify the right input directory with `-input`
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

Process multiple directories:

<pre><code># Process files from multiple source directories
python dumper.py -input "./2024_data" -output "./processed/2024"
python dumper.py -input "./2025_data" -output "./processed/2025"

# Process newest file from each subdirectory
for dir in data/*/; do python dumper.py -input "$dir" -output "./processed/$(basename $dir)"; done
</code></pre>

### Directory-Based Workflows

Organize processing by separating input and output:

<pre><code># Production workflow
python dumper.py -input "/data/incoming" -output "/data/processed" -no-hide

# Development workflow  
python dumper.py -input "./test_data" -output "./test_results"
</code></pre>

### Automation

The script can be integrated into automated workflows:

<pre><code># Daily report processing from specific directory
python dumper.py -input "/data/daily_reports" -output "./exports/$(date +%Y-%m-%d)"

# Weekly batch processing
python dumper.py -input "/data/weekly" -output "/archive/weekly"
</code></pre>

### Large Files

For very large Excel files:
- The script processes sheets one at a time to manage memory usage
- Progress is shown for each sheet being processed
- Consider using the `-no-hide` option to skip unnecessary hidden sheets
