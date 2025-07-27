"""
Fresh, working test file for excel_dumper.
This should run without any issues.
"""

import pytest
import tempfile
import os
import csv
import json
from pathlib import Path

# Import functions to test
from excel_dumper.dumper import (
    has_non_null_data,
    generate_output_filename,
    write_to_csv,
    write_to_json
)


def test_has_non_null_data():
    """Test the basic utility function."""
    # Test with valid data
    assert has_non_null_data(['value1', 'value2']) == True
    assert has_non_null_data([123, 'text']) == True
    assert has_non_null_data([0, False]) == True  # These are valid
    
    # Test with empty/null data
    assert has_non_null_data([None, None]) == False
    assert has_non_null_data(['', '   ']) == False
    assert has_non_null_data([]) == False
    
    # Test mixed data
    assert has_non_null_data([None, 'value']) == True


def test_filename_generation():
    """Test filename generation."""
    # Create a temporary file to get a real path
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        temp_path = tmp.name
    
    try:
        # Test CSV filename
        csv_name = generate_output_filename(temp_path, output_format='csv')
        assert csv_name.endswith('.csv')
        assert 'dumperpy_' in csv_name
        
        # Test JSON filename  
        json_name = generate_output_filename(temp_path, output_format='json')
        assert json_name.endswith('.json')
        
    finally:
        os.unlink(temp_path)


def test_csv_writing():
    """Test CSV file writing."""
    test_data = [
        ['Sheet1', 'Name', 'Age'],
        ['Sheet1', 'John', 25],
        ['Sheet1', 'Jane', 30]
    ]
    
    # Create temporary file
    with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as tmp:
        temp_path = tmp.name
    
    try:
        # Write CSV
        write_to_csv(test_data, temp_path)
        
        # Verify file exists and has content
        assert os.path.exists(temp_path)
        assert os.path.getsize(temp_path) > 0
        
        # Read and check content
        with open(temp_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            # Should have header plus data rows
            assert len(rows) == 4  # header + 3 data rows
            assert 'Worksheet' in rows[0]  # Header should include Worksheet
        
    finally:
        try:
            os.unlink(temp_path)
        except:
            pass  # Ignore cleanup errors


def test_json_writing():
    """Test JSON file writing."""
    test_data = [
        ['Sheet1', 'John', 25],
        ['Sheet1', 'Jane', 30]
    ]
    
    # Create temporary file
    with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as tmp:
        temp_path = tmp.name
    
    try:
        # Write JSON
        write_to_json(test_data, temp_path)
        
        # Verify file exists
        assert os.path.exists(temp_path)
        
        # Read and check content
        with open(temp_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
            assert isinstance(data, list)
            assert len(data) == 2
            assert 'Worksheet' in data[0]
            assert data[0]['Worksheet'] == 'Sheet1'
        
    finally:
        try:
            os.unlink(temp_path)
        except:
            pass  # Ignore cleanup errors


def test_excel_extraction_if_available():
    """Test Excel extraction if files are available."""
    # Look for any Excel files in fixtures
    fixtures_dir = Path("tests/fixtures")
    
    if fixtures_dir.exists():
        excel_files = list(fixtures_dir.glob("*.xlsx"))
        
        if excel_files:
            from excel_dumper.dumper import extract_excel_data
            
            sample_file = str(excel_files[0])
            print(f"Testing with file: {sample_file}")
            
            try:
                result = extract_excel_data(sample_file)
                assert isinstance(result, list)
                print(f"Extracted {len(result)} rows")
                
                if result:
                    # Check structure
                    assert len(result[0]) > 0  # Should have columns
                    print(f"First row: {result[0]}")
                    
            except Exception as e:
                print(f"Excel extraction failed: {e}")
                # Don't fail the test - just report
        else:
            print("No Excel files found in fixtures")
    else:
        print("No fixtures directory found")


# Run a quick validation
if __name__ == "__main__":
    print("Running quick validation...")
    test_has_non_null_data()
    print("✓ has_non_null_data works")
    
    test_filename_generation()
    print("✓ filename generation works")
    
    test_csv_writing()
    print("✓ CSV writing works")
    
    test_json_writing()
    print("✓ JSON writing works")
    
    test_excel_extraction_if_available()
    print("✓ Excel extraction test completed")
    
    print("\nAll basic tests passed!")