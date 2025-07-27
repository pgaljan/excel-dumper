"""
Example usage tests for excel_dumper.
"""

import pytest
import tempfile
from pathlib import Path

def test_quick_validation():
    """Quick test to verify the testing framework works."""
    from excel_dumper.dumper import has_non_null_data
    
    assert has_non_null_data(['value']) == True
    assert has_non_null_data([None, None]) == False
    print("✓ Quick validation passed!")

def test_import_works():
    """Test that we can import the excel_dumper module."""
    try:
        from excel_dumper.dumper import extract_excel_data, has_non_null_data
        assert extract_excel_data is not None
        assert has_non_null_data is not None
        print("✓ Imports working correctly")
    except ImportError as e:
        pytest.fail(f"Cannot import excel_dumper: {e}")

def test_simple_extraction(sample_xlsx_file):
    """Simple Excel extraction test."""
    try:
        from excel_dumper.dumper import extract_excel_data
        
        result = extract_excel_data(sample_xlsx_file)
        assert len(result) >= 1
        assert any('Employees' in str(row) or 'Summary' in str(row) for row in result)
        print("✓ Simple extraction test passed!")
            
    except Exception as e:
        pytest.fail(f"Simple extraction failed: {e}")

if __name__ == "__main__":
    print("Running example tests...")
    test_quick_validation()
    test_import_works()
    print("✓ Example tests completed!")
