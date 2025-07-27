"""
Core Excel extraction tests for excel_dumper.
"""

import pytest
import tempfile
from pathlib import Path
import openpyxl
from openpyxl import Workbook

# Import the function we're testing
from excel_dumper.dumper import extract_excel_data, has_non_null_data


class TestExcelExtraction:
    """Test core Excel data extraction functionality."""
    
    def test_extract_basic_xlsx(self, sample_xlsx_file):
        """Test basic Excel extraction functionality with XLSX file."""
        result = extract_excel_data(sample_xlsx_file)
        
        # Verify we got data back
        assert isinstance(result, list)
        assert len(result) > 0
        
        # Verify structure - each row should start with worksheet name
        for row in result:
            assert isinstance(row, list)
            assert len(row) > 0
            assert isinstance(row[0], str)  # First column should be worksheet name
        
        # Check that we have data from multiple sheets
        worksheet_names = {row[0] for row in result}
        assert len(worksheet_names) >= 2  # Should have at least 2 worksheets
        
        print(f"✓ Basic extraction test passed with {len(result)} rows from {len(worksheet_names)} sheets")


    def test_extract_with_row_numbers(self, sample_xlsx_file):
        """Test extraction with row numbers included."""
        result = extract_excel_data(sample_xlsx_file, include_row_numbers=True)
        
        assert len(result) > 0
        
        # With row numbers, structure should be: [worksheet, row_number, ...data]
        for row in result:
            assert len(row) >= 2
            assert isinstance(row[0], str)  # Worksheet name
            assert isinstance(row[1], int)  # Row number
            assert row[1] > 0  # Row numbers should be positive
        
        print(f"✓ Row numbers test passed with {len(result)} rows")


    def test_extract_without_row_numbers(self, sample_xlsx_file):
        """Test extraction without row numbers (default behavior)."""
        result = extract_excel_data(sample_xlsx_file, include_row_numbers=False)
        
        assert len(result) > 0
        
        # Without row numbers, structure should be: [worksheet, ...data]
        for row in result:
            assert isinstance(row[0], str)  # Worksheet name
        
        print(f"✓ No row numbers test passed with {len(result)} rows")


    def test_extract_formula_file_basic(self, formulas_xlsx_file):
        """Test basic extraction from formula file."""
        # Test with formulas enabled
        result_formulas = extract_excel_data(formulas_xlsx_file, include_formulas=True)
        assert len(result_formulas) > 0, "Should extract data with formulas enabled"
        
        # Test with formulas disabled
        result_values = extract_excel_data(formulas_xlsx_file, include_formulas=False)
        assert len(result_values) > 0, "Should extract data with formulas disabled"
        
        # Basic structure validation
        for row in result_formulas:
            assert len(row) > 0
            assert isinstance(row[0], str)  # Worksheet name
        
        print(f"✓ Formula file extraction: {len(result_formulas)} rows (with formulas), {len(result_values)} rows (calculated)")


    def test_extract_hidden_sheets_included(self, hidden_sheets_file):
        """Test extraction including hidden sheets."""
        result = extract_excel_data(hidden_sheets_file, include_hidden=True)
        
        worksheet_names = {row[0] for row in result}
        
        # Should include both visible and hidden sheets
        assert 'Visible' in worksheet_names
        assert 'Hidden' in worksheet_names
        
        # Verify we have data from hidden sheet
        hidden_rows = [row for row in result if row[0] == 'Hidden']
        assert len(hidden_rows) > 0
        
        print(f"✓ Hidden sheets included: {worksheet_names}")


    def test_extract_hidden_sheets_excluded(self, hidden_sheets_file):
        """Test extraction excluding hidden sheets."""
        result = extract_excel_data(hidden_sheets_file, include_hidden=False)
        
        worksheet_names = {row[0] for row in result}
        
        # Should only include visible sheets
        assert 'Visible' in worksheet_names
        assert 'Hidden' not in worksheet_names
        
        print(f"✓ Hidden sheets excluded: {worksheet_names}")


    def test_extract_empty_file(self, empty_xlsx_file):
        """Test handling of empty Excel files."""
        result = extract_excel_data(empty_xlsx_file)
        
        # Empty file should return empty list or list with no meaningful data
        assert isinstance(result, list)
        
        # Filter out any rows that might be completely empty
        non_empty_rows = [row for row in result if len(row) > 1 and has_non_null_data(row[1:])]
        assert len(non_empty_rows) == 0
        
        print("✓ Empty file handled correctly")


    def test_extract_nonexistent_file(self):
        """Test behavior when Excel file doesn't exist."""
        with pytest.raises(Exception) as exc_info:
            extract_excel_data('nonexistent_file.xlsx')
        
        # Should raise an exception mentioning the file
        assert 'nonexistent_file.xlsx' in str(exc_info.value)
        
        print("✓ Nonexistent file error handled correctly")


    def test_extract_preserves_worksheet_names(self, sample_xlsx_file):
        """Test that worksheet names are preserved correctly."""
        result = extract_excel_data(sample_xlsx_file)
        
        worksheet_names = {row[0] for row in result}
        
        # Should have expected worksheet names from our test file
        expected_names = {'Employees', 'Summary', 'Mixed Data'}
        
        # Check that we have at least some of the expected names
        found_names = worksheet_names.intersection(expected_names)
        assert len(found_names) > 0, f"Expected some of {expected_names}, found {worksheet_names}"
        
        # All worksheet names should be strings
        assert all(isinstance(name, str) for name in worksheet_names)
        assert all(len(name) > 0 for name in worksheet_names)
        
        print(f"✓ Worksheet names preserved: {worksheet_names}")


    def test_extract_handles_empty_rows(self, sample_xlsx_file):
        """Test that empty rows are properly filtered out."""
        result = extract_excel_data(sample_xlsx_file)
        
        # All returned rows should have non-null data
        for row in result:
            data_part = row[1:]  # Remove worksheet name
            assert has_non_null_data(data_part), f"Row {row} should have been filtered out"
        
        print(f"✓ Empty rows filtered correctly from {len(result)} rows")


    def test_extract_with_all_options(self, sample_xlsx_file):
        """Test extraction with all options enabled."""
        result = extract_excel_data(
            sample_xlsx_file,
            include_hidden=True,
            include_row_numbers=True,
            include_formulas=True
        )
        
        assert len(result) > 0
        
        # With all options, structure should be: [worksheet, row_number, ...data]
        for row in result:
            assert len(row) >= 2
            assert isinstance(row[0], str)  # Worksheet name
            assert isinstance(row[1], int)  # Row number
        
        print(f"✓ All options test passed with {len(result)} rows")


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
    
    print("✓ has_non_null_data utility function works correctly")


def test_extract_excel_data_function_signature():
    """Test that the extract_excel_data function has the expected signature."""
    import inspect
    
    sig = inspect.signature(extract_excel_data)
    params = list(sig.parameters.keys())
    
    # Check expected parameters exist
    assert 'filename' in params
    assert 'include_hidden' in params
    assert 'include_row_numbers' in params
    assert 'include_formulas' in params
    
    # Check default values
    assert sig.parameters['include_hidden'].default == True
    assert sig.parameters['include_row_numbers'].default == False
    assert sig.parameters['include_formulas'].default == False
    
    print("✓ Function signature is correct")
