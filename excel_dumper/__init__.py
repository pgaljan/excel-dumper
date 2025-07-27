"""
Excel Dumper - Cross-platform Excel ETL preprocessor

A comprehensive tool for extracting data from Excel worksheets and converting them
to CSV format for data pipeline ingestion, reporting, and auditing purposes.

This package provides functionality to:
- Extract data from all Excel formats (.xlsx, .xls, .xlsm, .xlsb)
- Process multiple worksheets within a single file
- Handle hidden worksheets with configurable options
- Export formulas or calculated values
- Include row number tracking for audit trails
- Generate timestamped output files with collision-safe naming
- Support cross-platform file path handling

Example usage:
    from excel_dumper import main, extract_excel_data, write_to_csv
    
    # Run the main CLI interface
    main()
    
    # Or use individual functions programmatically
    data = extract_excel_data('data.xlsx', include_formulas=True)
    write_to_csv(data, 'output.csv', include_row_numbers=True)

Command-line usage:
    dumper -file data.xlsx -rownumbers -formulas
    excel-dumper -input ./data -output ./processed
"""

__version__ = "1.0.0"
__author__ = "pgaljan"
__email__ = "galjan@gmail.com"
__description__ = "Cross-platform Excel ETL preprocessor for data pipeline ingestion and auditing"
__url__ = "https://github.com/pgaljan/excel-dumper"

# Import main functions for programmatic use
from .dumper import (
    main,
    extract_excel_data,
    write_to_csv,
    write_to_json,
    find_newest_excel_file,
    generate_output_filename,
    has_non_null_data,
    show_help
)

# Define what gets imported with "from excel_dumper import *"
__all__ = [
    # Main entry point
    "main",
    
    # Core data processing functions
    "extract_excel_data",
    "write_to_csv", 
    "write_to_json",
    
    # Utility functions
    "find_newest_excel_file",
    "generate_output_filename",
    "has_non_null_data",
    "show_help",
    
    # Package metadata
    "__version__",
    "__author__",
    "__email__",
    "__description__",
    "__url__",
]

# Package-level configuration
DEFAULT_EXCEL_EXTENSIONS = ['.xlsx', '.xls', '.xlsm', '.xlsb']
DEFAULT_OUTPUT_FORMAT = 'csv'
DEFAULT_INCLUDE_HIDDEN = True
DEFAULT_INCLUDE_ROW_NUMBERS = False
DEFAULT_INCLUDE_FORMULAS = False

# Convenience functions for common use cases
def quick_extract(filename, output_format='csv', include_row_numbers=False):
    """
    Quick extraction function for simple use cases.
    
    Args:
        filename (str): Path to Excel file
        output_format (str): 'csv' or 'json'
        include_row_numbers (bool): Include row numbers in output
    
    Returns:
        str: Path to generated output file
    """
    from pathlib import Path
    
    if not Path(filename).exists():
        raise FileNotFoundError(f"Excel file not found: {filename}")
    
    # Extract data
    data = extract_excel_data(filename, include_row_numbers=include_row_numbers)
    
    if not data:
        raise ValueError("No data found to export")
    
    # Generate output filename
    output_file = generate_output_filename(filename, output_format=output_format)
    
    # Write output
    if output_format == 'json':
        write_to_json(data, output_file, include_row_numbers)
    else:
        write_to_csv(data, output_file, include_row_numbers)
    
    return output_file


def batch_process(input_dir=".", output_dir=None, output_format='csv'):
    """
    Process all Excel files in a directory.
    
    Args:
        input_dir (str): Directory to search for Excel files
        output_dir (str): Output directory (default: same as input)
        output_format (str): 'csv' or 'json'
    
    Returns:
        list: Paths to generated output files
    """
    from pathlib import Path
    import glob
    
    search_path = Path(input_dir)
    output_files = []
    
    # Find all Excel files
    excel_files = []
    for ext in DEFAULT_EXCEL_EXTENSIONS:
        excel_files.extend(search_path.glob(f"*{ext}"))
    
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in directory: {input_dir}")
    
    # Process each file
    for excel_file in excel_files:
        try:
            data = extract_excel_data(str(excel_file))
            if data:
                output_file = generate_output_filename(
                    str(excel_file), 
                    output_dir, 
                    output_format
                )
                
                if output_format == 'json':
                    write_to_json(data, output_file)
                else:
                    write_to_csv(data, output_file)
                
                output_files.append(output_file)
                
        except Exception as e:
            print(f"Warning: Could not process {excel_file}: {e}")
            continue
    
    return output_files


# Version check function
def check_dependencies():
    """
    Check if all required dependencies are available.
    
    Returns:
        dict: Status of each dependency
    """
    dependencies = {}
    
    try:
        import pandas
        dependencies['pandas'] = {
            'available': True,
            'version': pandas.__version__
        }
    except ImportError:
        dependencies['pandas'] = {
            'available': False,
            'version': None
        }
    
    try:
        import openpyxl
        dependencies['openpyxl'] = {
            'available': True,
            'version': openpyxl.__version__
        }
    except ImportError:
        dependencies['openpyxl'] = {
            'available': False,
            'version': None
        }
    
    try:
        import xlrd
        dependencies['xlrd'] = {
            'available': True,
            'version': xlrd.__version__
        }
    except ImportError:
        dependencies['xlrd'] = {
            'available': False,
            'version': None
        }
    
    return dependencies


# Add convenience functions to __all__
__all__.extend([
    'quick_extract',
    'batch_process', 
    'check_dependencies',
    'DEFAULT_EXCEL_EXTENSIONS',
    'DEFAULT_OUTPUT_FORMAT',
    'DEFAULT_INCLUDE_HIDDEN',
    'DEFAULT_INCLUDE_ROW_NUMBERS',
    'DEFAULT_INCLUDE_FORMULAS',
])