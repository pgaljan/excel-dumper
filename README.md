# Dumper

Cross-platform Excel ETL pre-processor that flattens Excel files into a predictable CSV format, preserving key source metadata for easy ingestion into data pipelines. Exports values or formulas, simplifying audit processes. Available in both **Python** and **PowerShell** versions.

## Features

- **Auto-discovery**: Automatically finds and processes the newest Excel file in the specified directory
- **Multi-sheet support**: Extracts data from all worksheets in an Excel file
- **Hidden sheet control**: Option to include or skip hidden worksheets
- **Smart filtering**: Only exports rows containing non-null data
- **Formula extraction**: Option to show formulas instead of calculated values (.xlsx/.xlsm only)
- **Worksheet identification**: Prepends worksheet name as the first column in output
- **Row number tracking**: Option to include Excel row numbers for data traceability
- **Collision-safe naming**: Automatically adds incremental numbers if output files already exist
- **Custom output directory**: Specify where to save the exported CSV files
- **Flexible input sources**: Search for files in any directory or process specific files
- **Multiple Excel formats**: Supports .xlsx, .xls, .xlsm, and .xlsb files
- **Timestamped output**: Uses file modification time for consistent naming (ISO 8601 format)

## Quick Start

### Python Version
```bash
# Install dependencies
pip install pandas openpyxl xlrd

# Basic usage
python dumper.py
python dumper.py -file "data.xlsx" -rownumbers
```

### PowerShell Version
```powershell
# Install dependencies
Install-Module ImportExcel -Force

# Basic usage
.\dumper.ps1
.\dumper.ps1 -File "data.xlsx" -RowNumbers
```

## Installation

### Python Version

**Prerequisites:** Python 3.6 or higher

**Install Dependencies:**
```bash
pip install pandas openpyxl xlrd
```

### PowerShell Version

**Prerequisites:** PowerShell 5.1 or PowerShell Core 6+

**Install Dependencies:**
*with admin rights*
```powershell
Install-Module ImportExcel -Force
```
*without admin rights*

```powershell
Install-Module ImportExcel -Force -Scope CurrentUser
```

## Command Line Options

| Feature | Python Version | PowerShell Version | Description |
|---------|----------------|-------------------|-------------|
| **Specify file** | `-file FILE` | `-File FILE` | Process specific Excel file |
| **Input directory** | `-input DIR` | `-InputDir DIR` | Directory to search for Excel files |
| **Output directory** | `-output DIR` | `-OutputDir DIR` | Directory for CSV output |
| **Skip hidden sheets** | `-no-hide` | `-NoHide` | Exclude hidden worksheets |
| **Include row numbers** | `-rownumbers` | `-RowNumbers` | Add Excel row numbers to output |
| **Show formulas** | `-formulas` | `-Formulas` | Show formulas instead of values (.xlsx/.xlsm only) |
| **Help** | `-help` | `-Help` | Display detailed help |

## Usage Examples

### Basic Usage

**Python:**
```bash
# Process newest Excel file from current directory
python dumper.py

# Process specific file
python dumper.py -file "data.xlsx"

# Include row numbers and formulas
python dumper.py -file "data.xlsx" -rownumbers -formulas
```

**PowerShell:**
```powershell
# Process newest Excel file from current directory
.\dumper.ps1

# Process specific file
.\dumper.ps1 -File "data.xlsx"

# Include row numbers and formulas
.\dumper.ps1 -File "data.xlsx" -RowNumbers -Formulas
```

### Advanced Usage

**Python:**
```bash
# Comprehensive processing with all options
python dumper.py -input "./source" -output "./exports" -file "report.xlsx" -no-hide -rownumbers -formulas

# Process from specific directory
python dumper.py -input "C:\Reports" -output "C:\Processed"
```

**PowerShell:**
```powershell
# Comprehensive processing with all options
.\dumper.ps1 -InputDir "./source" -OutputDir "./exports" -File "report.xlsx" -NoHide -RowNumbers -Formulas

# Process from specific directory
.\dumper.ps1 -InputDir "C:\Reports" -OutputDir "C:\Processed"
```

### Get Help

**Python:**
```bash
python dumper.py -help
```

**PowerShell:**
```powershell
.\dumper.ps1 -Help
```

## Output Format

### File Naming Convention

Both versions create files with the pattern:
```
dumper[py|ps]_[original_filename]_[timestamp].csv
```

- **Python version**: `dumperpy_Sales_Report_2025-07-21T14-30-52-0500.csv`
- **PowerShell version**: `dumperps_Sales_Report_2025-07-21T14-30-52-0500.csv`
- **Timestamp format**: ISO 8601 with colons replaced by hyphens
- **Collision handling**: Adds `(1)`, `(2)`, etc. if file exists

### CSV Structure

The output CSV contains:

1. **First column**: Worksheet name
2. **Second column**: Excel row number (if row numbers option used)
3. **Remaining columns**: Original data from Excel worksheets (`Column_1`, `Column_2`, etc.)

**Without row numbers:**
```csv
Worksheet,Column_1,Column_2,Column_3
Sheet1,John Doe,Sales Manager,50000
Sheet1,Jane Smith,Developer,65000
Summary,Total Employees,,2
```

**With row numbers:**
```csv
Worksheet,Row_Number,Column_1,Column_2,Column_3
Sheet1,2,John Doe,Sales Manager,50000
Sheet1,3,Jane Smith,Developer,65000
Summary,5,Total Employees,,2
```

**With formulas (when `-formulas` option used):**
```csv
Worksheet,Column_1,Column_2,Column_3
Sheet1,John Doe,Sales Manager,50000
Summary,Total Employees,"FORMULA: =COUNTA(A:A)-1",2
Summary,Average Salary,"FORMULA: =AVERAGE(C2:C3)",57500
```

## Supported Excel Formats

Both versions support:
- **.xlsx** - Excel 2007+ (Open XML Format)
- **.xls** - Excel 97-2003 (Binary Format)
- **.xlsm** - Excel Macro-Enabled Workbook
- **.xlsb** - Excel Binary Workbook

**Note:** Formula extraction (`-formulas`/`-Formulas`) only works with .xlsx and .xlsm files.

## Dependencies

### Python Version
- **pandas** - Data manipulation and Excel reading
- **openpyxl** - Excel 2007+ file support and formula extraction
- **xlrd** - Legacy Excel (.xls) file support
- Standard library modules: `argparse`, `csv`, `glob`, `os`, `sys`, `datetime`, `pathlib`

### PowerShell Version
- **ImportExcel** module - Excel file processing
- PowerShell 5.1 or PowerShell Core 6+

## Platform Compatibility

### Python Version
- **Windows, macOS, Linux** - Full cross-platform support
- Uses `pathlib` for platform-independent path handling

### PowerShell Version
- **Windows** - Native PowerShell 5.1+
- **macOS, Linux** - PowerShell Core 6+
- Cross-platform path handling with PowerShell cmdlets

## Error Handling

Both versions provide clear error messages for:
- **File not found**: When specified file doesn't exist
- **No Excel files**: When no Excel files found in directory
- **Missing dependencies**: When required libraries/modules aren't installed
- **Read errors**: When Excel file is corrupted or inaccessible
- **Write errors**: When output location is not writable
- **Directory creation**: Automatically creates output directories

## Example Workflows

### Daily Report Processing

**Python:**
```bash
# Process daily reports
python dumper.py -input "/data/daily_reports" -output "./exports/$(date +%Y-%m-%d)"
```

**PowerShell:**
```powershell
# Process daily reports
.\dumper.ps1 -InputDir "C:\Data\DailyReports" -OutputDir "C:\Exports\$(Get-Date -Format 'yyyy-MM-dd')"
```

### Batch Processing Multiple Directories

**Python:**
```bash
# Process multiple directories
for dir in data/*/; do 
    python dumper.py -input "$dir" -output "./processed/$(basename $dir)"
done
```

**PowerShell:**
```powershell
# Process multiple directories
Get-ChildItem "C:\Data" -Directory | ForEach-Object {
    .\dumper.ps1 -InputDir $_.FullName -OutputDir "C:\Processed\$($_.Name)"
}
```

## Troubleshooting

### Common Issues

**Python - "Error: Required libraries not found"**
```bash
pip install pandas openpyxl xlrd
python -c "import pandas, openpyxl; print('Dependencies OK')"
```

**PowerShell - "ImportExcel module not found"**
```powershell
Install-Module ImportExcel -Force -Scope CurrentUser
Import-Module ImportExcel
```

**"No Excel files found in directory"**
- Verify directory path with `-input`/`-InputDir`
- Check file extensions (.xlsx, .xls, .xlsm, .xlsb)
- Use `-file`/`-File` to specify exact filename

**"Permission denied" errors**
- Check write permissions for output directory
- Close Excel files that might be locking files
- Run with appropriate privileges

**"Error reading Excel file"**
- Verify file isn't corrupted by opening in Excel
- Check file isn't password-protected
- Ensure file isn't in use by another application

## Choosing Between Versions

### Use Python Version When:
- You need maximum cross-platform compatibility
- You're integrating with existing Python data pipelines
- You prefer pip-based dependency management
- You need the most mature Excel processing capabilities

### Use PowerShell Version When:
- You're working primarily in Windows environments
- You prefer PowerShell-based automation workflows
- You want native integration with Windows administration tasks
- You're already using PowerShell for other data processing

## Advanced Configuration

### Large File Processing
Both versions handle large files efficiently by:
- Processing sheets one at a time to manage memory
- Showing progress for each sheet
- Using streaming approaches where possible

### Integration with CI/CD
Both scripts can be integrated into automated workflows:

**Python in GitHub Actions:**
```yaml
- name: Process Excel files
  run: |
    pip install pandas openpyxl xlrd
    python dumper.py -input "./data" -output "./processed"
```

**PowerShell in Azure DevOps:**
```yaml
- task: PowerShell@2
  inputs:
    targetType: 'inline'
    script: |
      Install-Module ImportExcel -Force
      .\dumper.ps1 -InputDir ".\data" -OutputDir ".\processed"
```

## Performance Considerations

- **Hidden sheets**: Use `-no-hide`/`-NoHide` to skip unnecessary processing
- **Formula extraction**: Only enable when needed as it requires additional processing
- **Large workbooks**: Both versions stream data to minimize memory usage
- **Network drives**: Local processing is faster than network-mounted drives
