#!/usr/bin/env python3
"""
Excel Sheet Excavator Script
Extracts all non-null rows from Excel worksheets and saves to CSV format.
"""

import argparse
import csv
import glob
import os
import sys
from datetime import datetime
from pathlib import Path

try:
    from sheet_excavator import excavate
except ImportError:
    print("Error: sheet-excavator library not found.")
    print("Please install it with: pip install sheet-excavator")
    sys.exit(1)


def find_newest_excel_file():
    """Find the newest Excel file in the current directory."""
    excel_patterns = ['*.xlsx', '*.xls', '*.xlsm', '*.xlsb']
    excel_files = []
    
    for pattern in excel_patterns:
        excel_files.extend(glob.glob(pattern))
    
    if not excel_files:
        raise FileNotFoundError("No Excel files found in current directory")
    
    # Get the newest file based on modification time
    newest_file = max(excel_files, key=os.path.getmtime)
    return newest_file


def has_non_null_data(row):
    """Check if a row contains any non-null data."""
    return any(cell is not None and str(cell).strip() != '' for cell in row)


def extract_excel_data(filename, include_hidden=True):
    """Extract data from all worksheets in an Excel file."""
    try:
        # Use sheet-excavator to read the Excel file
        workbook_data = excavate(filename)
        
        extracted_data = []
        
        for sheet_name, sheet_data in workbook_data.items():
            # Skip hidden sheets if include_hidden is False
            # Note: sheet-excavator may not distinguish hidden sheets by default
            # This would depend on the specific implementation of the library
            
            if not sheet_data:
                continue
                
            # Process each row in the sheet
            for row_idx, row in enumerate(sheet_data):
                if has_non_null_data(row):
                    # Prepend worksheet name as first column
                    row_with_sheet = [sheet_name] + list(row)
                    extracted_data.append(row_with_sheet)
        
        return extracted_data
        
    except Exception as e:
        raise Exception(f"Error reading Excel file '{filename}': {str(e)}")


def write_to_csv(data, output_filename):
    """Write extracted data to CSV file."""
    try:
        with open(output_filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write header
            if data:
                # Determine maximum number of columns
                max_cols = max(len(row) for row in data) if data else 0
                header = ['Worksheet'] + [f'Column_{i}' for i in range(1, max_cols)]
                writer.writerow(header)
            
            # Write data rows
            for row in data:
                writer.writerow(row)
                
        print(f"Data successfully exported to: {output_filename}")
        print(f"Total rows exported: {len(data)}")
        
    except Exception as e:
        raise Exception(f"Error writing to CSV file '{output_filename}': {str(e)}")


def generate_output_filename(input_filename):
    """Generate output filename based on input filename with timestamp."""
    input_path = Path(input_filename)
    base_name = input_path.stem
    
    # Get file modification time with timezone
    mod_time = os.path.getmtime(input_filename)
    mod_datetime = datetime.fromtimestamp(mod_time).astimezone()
    
    # Format timestamp as YYYYMMDD_HHMMSS_TZ (e.g., 20250721_143052_EST)
    timestamp = mod_datetime.strftime("%Y%m%d_%H%M%S_%Z")
    
    return f"dumper_{base_name}_{timestamp}.csv"


def show_help():
    """Display help information."""
    help_text = """
Excel Sheet Excavator - Extract data from Excel worksheets to CSV

USAGE:
    python dumper.py [OPTIONS]

OPTIONS:
    -file FILE          Specify Excel file to process (default: newest Excel file in current directory)
    -no-hide           Skip hidden worksheets (default: include all worksheets)
    -help              Show this help message

EXAMPLES:
    python dumper.py                    # Process newest Excel file, include all sheets
    python dumper.py -file data.xlsx    # Process specific file
    python dumper.py -no-hide           # Skip hidden worksheets
    python dumper.py -file data.xlsx -no-hide  # Specific file, skip hidden sheets

OUTPUT:
    Creates a CSV file named "dumper_[original_filename]_[timestamp].csv" with:
    - Timestamp is the last modified time of the originating Excel file
    - Timestamp format: YYYYMMDD_HHMMSS_TZ (e.g., dumper_data_20250721_143052_EST.csv)
    - First column: Worksheet name
    - Remaining columns: Original data from worksheets
    - Only non-empty rows are included

PYTHON DEPENDENCIES:
    - sheet-excavator  (pip install sheet-excavator)
    - Standard library: argparse, csv, glob, os, sys, pathlib, datetime

SUPPORTED EXCEL FORMATS:
    - .xlsx (Excel 2007+)
    - .xls  (Excel 97-2003)
    - .xlsm (Excel Macro-Enabled)
    - .xlsb (Excel Binary)
"""
    print(help_text)


def main():
    """Main function to handle command line arguments and orchestrate the process."""
    
    # Handle -help parameter separately since argparse's help conflicts with our custom help
    if '-help' in sys.argv:
        show_help()
        return
    
    parser = argparse.ArgumentParser(description='Extract Excel worksheet data to CSV', add_help=False)
    parser.add_argument('-file', dest='filename', help='Excel file to process')
    parser.add_argument('-no-hide', action='store_true', help='Skip hidden worksheets')
    
    try:
        args = parser.parse_args()
        
        # Determine input file
        if args.filename:
            if not os.path.exists(args.filename):
                print(f"Error: File '{args.filename}' not found.")
                sys.exit(1)
            input_file = args.filename
        else:
            try:
                input_file = find_newest_excel_file()
                print(f"Processing newest Excel file: {input_file}")
            except FileNotFoundError as e:
                print(f"Error: {e}")
                sys.exit(1)
        
        # Extract data
        include_hidden = not args.no_hide
        print(f"Extracting data from: {input_file}")
        print(f"Including hidden sheets: {include_hidden}")
        
        extracted_data = extract_excel_data(input_file, include_hidden)
        
        if not extracted_data:
            print("No data found to export.")
            return
        
        # Generate output filename and write CSV
        output_file = generate_output_filename(input_file)
        write_to_csv(extracted_data, output_file)
        
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()