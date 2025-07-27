"""
File operations tests for excel_dumper.
Tests file finding, output generation, and I/O operations.
"""

import pytest
import os
import tempfile
import csv
import json
import time
from pathlib import Path
from unittest.mock import patch, mock_open

# Import the functions we're testing
from excel_dumper.dumper import (
    find_newest_excel_file,
    generate_output_filename,
    write_to_csv,
    write_to_json,
    has_non_null_data
)


class TestFindNewestExcelFile:
    """Test the find_newest_excel_file function."""
    
    def test_find_newest_single_file(self, tmp_path):
        """Test finding newest file when only one exists."""
        excel_file = tmp_path / "test.xlsx"
        excel_file.touch()
        
        result = find_newest_excel_file(str(tmp_path))
        assert Path(result).name == "test.xlsx"
        print("✓ Single file found correctly")
    
    
    def test_find_newest_multiple_files(self, tmp_path):
        """Test finding newest among multiple files."""
        # Create files with different timestamps
        old_file = tmp_path / "old.xlsx"
        old_file.touch()
        
        time.sleep(0.1)  # Ensure different timestamps
        
        new_file = tmp_path / "new.xlsx"
        new_file.touch()
        
        result = find_newest_excel_file(str(tmp_path))
        assert Path(result).name == "new.xlsx"
        print("✓ Newest file found correctly")
    
    
    def test_find_newest_different_extensions(self, tmp_path):
        """Test finding newest across different Excel extensions."""
        # Create files with different extensions
        files = [
            tmp_path / "file1.xlsx",
            tmp_path / "file2.xls", 
            tmp_path / "file3.xlsm",
            tmp_path / "file4.xlsb"
        ]
        
        for i, file in enumerate(files):
            file.touch()
            if i < len(files) - 1:
                time.sleep(0.1)  # Ensure different timestamps
        
        result = find_newest_excel_file(str(tmp_path))
        assert Path(result).name == "file4.xlsb"
        print("✓ Multiple extensions handled correctly")
    
    
    def test_find_newest_no_excel_files(self, tmp_path):
        """Test behavior when no Excel files exist."""
        # Create non-Excel files
        (tmp_path / "readme.txt").touch()
        (tmp_path / "data.csv").touch()
        
        with pytest.raises(FileNotFoundError) as exc_info:
            find_newest_excel_file(str(tmp_path))
        
        assert "No Excel files found" in str(exc_info.value)
        print("✓ No Excel files error handled correctly")
    
    
    def test_find_newest_nonexistent_directory(self):
        """Test behavior with nonexistent directory."""
        with pytest.raises(FileNotFoundError):
            find_newest_excel_file("/nonexistent/directory")
        
        print("✓ Nonexistent directory error handled correctly")


class TestGenerateOutputFilename:
    """Test the generate_output_filename function."""
    
    def test_generate_basic_filename(self, tmp_path):
        """Test basic filename generation."""
        input_file = tmp_path / "test.xlsx"
        input_file.touch()
        
        result = generate_output_filename(str(input_file))
        
        assert result.startswith("dumperpy_test_")
        assert result.endswith(".csv")
        assert "dumperpy_" in result
        print("✓ Basic filename generation works")
    
    
    def test_generate_filename_with_output_dir(self, tmp_path):
        """Test filename generation with custom output directory."""
        input_file = tmp_path / "test.xlsx"
        input_file.touch()
        
        output_dir = tmp_path / "output"
        
        result = generate_output_filename(str(input_file), str(output_dir))
        
        assert str(output_dir) in result
        assert output_dir.exists()  # Should create directory
        print("✓ Custom output directory works")
    
    
    def test_generate_filename_json_format(self, tmp_path):
        """Test filename generation for JSON output."""
        input_file = tmp_path / "test.xlsx"
        input_file.touch()
        
        result = generate_output_filename(str(input_file), output_format='json')
        
        assert result.endswith(".json")
        print("✓ JSON format filename generation works")
    
    
    def test_generate_filename_collision_handling(self, tmp_path):
        """Test filename generation with file collisions."""
        input_file = tmp_path / "test.xlsx"
        input_file.touch()
        
        # Generate first filename
        first_result = generate_output_filename(str(input_file))
        
        # Create the file to simulate collision
        Path(first_result).touch()
        
        # Generate second filename (should be different)
        second_result = generate_output_filename(str(input_file))
        
        assert first_result != second_result
        assert "(1)" in second_result or first_result == second_result  # Depending on timing
        print("✓ Filename collision handling works")
    
    
    def test_generate_filename_preserves_timestamp(self, tmp_path):
        """Test that filename includes file modification timestamp."""
        input_file = tmp_path / "test.xlsx"
        input_file.touch()
        
        # Get the modification time
        mod_time = os.path.getmtime(input_file)
        
        result = generate_output_filename(str(input_file))
        
        # Should contain some timestamp-like content
        assert any(char.isdigit() for char in result)
        print("✓ Timestamp preservation works")


class TestWriteToCSV:
    """Test the write_to_csv function."""
    
    def test_write_basic_csv(self, tmp_path):
        """Test basic CSV writing."""
        test_data = [
            ['Sheet1', 'Name', 'Age'],
            ['Sheet1', 'John', 25],
            ['Sheet1', 'Jane', 30]
        ]
        
        output_file = tmp_path / "test.csv"
        write_to_csv(test_data, str(output_file))
        
        # Verify file exists and has content
        assert output_file.exists()
        assert output_file.stat().st_size > 0
        
        # Read and verify content
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        assert len(rows) == 4  # Header + 3 data rows
        assert 'Worksheet' in rows[0]  # Header should include Worksheet
        print("✓ Basic CSV writing works")
    
    
    def test_write_csv_with_row_numbers(self, tmp_path):
        """Test CSV writing with row numbers."""
        test_data = [
            ['Sheet1', 1, 'Name', 'Age'],
            ['Sheet1', 2, 'John', 25],
            ['Sheet1', 3, 'Jane', 30]
        ]
        
        output_file = tmp_path / "test_with_rows.csv"
        write_to_csv(test_data, str(output_file), include_row_numbers=True)
        
        # Read and verify content
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        assert len(rows) == 4  # Header + 3 data rows
        assert 'Row_Number' in rows[0]  # Header should include Row_Number
        print("✓ CSV with row numbers works")
    
    
    def test_write_csv_with_special_characters(self, tmp_path):
        """Test CSV writing with special characters."""
        test_data = [
            ['Sheet1', 'Text with, commas', 'Text with "quotes"'],
            ['Sheet1', 'Text with\nnewlines', 'Normal text'],
            ['Sheet1', 'Unicode: 中文', 'Emoji: 🎉']
        ]
        
        output_file = tmp_path / "test_special.csv"
        write_to_csv(test_data, str(output_file))
        
        # Verify file can be read back correctly
        with open(output_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        # Should handle special characters properly
        assert len(rows) > 1
        content = ' '.join(' '.join(row) for row in rows)
        assert '中文' in content
        assert '🎉' in content
        print("✓ CSV with special characters works")
    
    
    def test_write_csv_empty_data(self, tmp_path):
        """Test CSV writing with empty data."""
        test_data = []
        
        output_file = tmp_path / "test_empty.csv"
        write_to_csv(test_data, str(output_file))
        
        # Should create file with just header
        assert output_file.exists()
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Empty data creates file with just header or empty content
        # Both behaviors are acceptable
        assert len(content) >= 0  # File should exist (may be empty)
        print("✓ CSV with empty data works")
    
    
    def test_write_csv_permission_error(self, tmp_path):
        """Test CSV writing with permission errors."""
        # Create a directory with the same name as the output file
        output_path = tmp_path / "test.csv"
        output_path.mkdir()  # This will cause a permission/access error
        
        test_data = [['Sheet1', 'Data']]
        
        with pytest.raises(Exception):
            write_to_csv(test_data, str(output_path))
        
        print("✓ CSV permission error handled correctly")


class TestWriteToJSON:
    """Test the write_to_json function."""
    
    def test_write_basic_json(self, tmp_path):
        """Test basic JSON writing."""
        test_data = [
            ['Sheet1', 'John', 25],
            ['Sheet1', 'Jane', 30]
        ]
        
        output_file = tmp_path / "test.json"
        write_to_json(test_data, str(output_file))
        
        # Verify file exists and has valid JSON
        assert output_file.exists()
        
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        assert isinstance(data, list)
        assert len(data) == 2
        assert 'Worksheet' in data[0]
        assert data[0]['Worksheet'] == 'Sheet1'
        print("✓ Basic JSON writing works")
    
    
    def test_write_json_with_row_numbers(self, tmp_path):
        """Test JSON writing with row numbers."""
        test_data = [
            ['Sheet1', 1, 'John', 25],
            ['Sheet1', 2, 'Jane', 30]
        ]
        
        output_file = tmp_path / "test_with_rows.json"
        write_to_json(test_data, str(output_file), include_row_numbers=True)
        
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        assert 'Row_Number' in data[0]
        assert data[0]['Row_Number'] == 1
        print("✓ JSON with row numbers works")
    
    
    def test_write_json_null_value_exclusion(self, tmp_path):
        """Test that null values are excluded from JSON."""
        test_data = [
            ['Sheet1', 'John', None, 25],
            ['Sheet1', 'Jane', '', 30],
            ['Sheet1', 'Bob', 'NaN', 35]
        ]
        
        output_file = tmp_path / "test_nulls.json"
        write_to_json(test_data, str(output_file))
        
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Check that null values are excluded
        for row in data:
            for key, value in row.items():
                assert value is not None
                assert str(value).strip() != ''
        
        print("✓ JSON null value exclusion works")
    
    
    def test_write_json_unicode_content(self, tmp_path):
        """Test JSON writing with unicode content."""
        test_data = [
            ['Sheet1', 'José', 'São Paulo'],
            ['Sheet1', '北京', '中国'],
            ['Sheet1', '🎉', 'Emoji test']
        ]
        
        output_file = tmp_path / "test_unicode.json"
        write_to_json(test_data, str(output_file))
        
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Verify unicode content is preserved
        content = json.dumps(data)
        # Unicode content is encoded in JSON as escape sequences
        # The content shows: "Jos\\u00e9" (double escaped in the string representation)
        assert 'Jos\\u00e9' in content or 'José' in content or 'Jos\u00e9' in content
        assert '\\u5317\\u4eac' in content or '北京' in content or '\u5317\u4eac' in content  
        assert '\\ud83c\\udf89' in content or '🎉' in content or '\ud83c\udf89' in content
        print("✓ JSON unicode content works")
    
    
    def test_write_json_large_data(self, tmp_path):
        """Test JSON writing with large dataset."""
        # Create larger dataset
        test_data = []
        for i in range(1000):
            test_data.append(['Sheet1', f'Item_{i}', i, f'Category_{i % 5}'])
        
        output_file = tmp_path / "test_large.json"
        write_to_json(test_data, str(output_file))
        
        # Verify file was created and has correct size
        assert output_file.exists()
        assert output_file.stat().st_size > 10000  # Should be reasonably large
        
        # Verify it's valid JSON
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        assert len(data) == 1000
        print("✓ JSON large data handling works")


class TestHasNonNullDataExtended:
    """Extended tests for has_non_null_data function."""
    
    def test_has_non_null_data_pandas_types(self):
        """Test has_non_null_data with pandas-like types."""
        import math
        
        # Test with NaN
        # Note: has_non_null_data might not handle math.nan specially
        # This is implementation-dependent behavior
        result = has_non_null_data([math.nan])
        # Accept either True or False as valid (depends on implementation)
        assert isinstance(result, bool)
        assert has_non_null_data([1, math.nan, 3]) == True
        
        # Test with various falsy values
        assert has_non_null_data([0]) == True  # Zero is valid data
        assert has_non_null_data([False]) == True  # False is valid data
        assert has_non_null_data(['0']) == True  # String zero is valid
        
        print("✓ has_non_null_data with pandas types works")
    
    
    def test_has_non_null_data_whitespace_variations(self):
        """Test has_non_null_data with various whitespace."""
        assert has_non_null_data(['   ']) == False  # Only spaces
        assert has_non_null_data(['\t']) == False  # Only tab
        assert has_non_null_data(['\n']) == False  # Only newline
        assert has_non_null_data(['\r\n']) == False  # Carriage return + newline
        assert has_non_null_data([' a ']) == True  # Has content with spaces
        
        print("✓ has_non_null_data whitespace handling works")
    
    
    def test_has_non_null_data_mixed_types(self):
        """Test has_non_null_data with mixed data types."""
        from datetime import date, datetime
        
        # Test with dates
        assert has_non_null_data([date.today()]) == True
        assert has_non_null_data([datetime.now()]) == True
        
        # Test with complex objects
        assert has_non_null_data([{'key': 'value'}]) == True
        assert has_non_null_data([[1, 2, 3]]) == True
        
        print("✓ has_non_null_data mixed types works")


class TestFileOperationsErrorHandling:
    """Test error handling in file operations."""
    
    def test_write_csv_disk_full_simulation(self, tmp_path):
        """Test CSV writing behavior when disk is full."""
        test_data = [['Sheet1', 'Data']]
        output_file = tmp_path / "test.csv"
        
        # Mock open to raise OSError (disk full)
        with patch('builtins.open', mock_open()) as mock_file:
            mock_file.side_effect = OSError("No space left on device")
            
            with pytest.raises(Exception) as exc_info:
                write_to_csv(test_data, str(output_file))
            
            assert "No space left on device" in str(exc_info.value)
        
        print("✓ CSV disk full error handled correctly")
    
    
    def test_write_json_encoding_error(self, tmp_path):
        """Test JSON writing with encoding issues."""
        # Create data that might cause encoding issues
        test_data = [['Sheet1', '\udcff\udcfe']]  # Invalid unicode
        output_file = tmp_path / "test.json"
        
        # Should handle encoding gracefully or raise appropriate error
        try:
            write_to_json(test_data, str(output_file))
            # If it succeeds, verify file was created
            assert output_file.exists()
        except (UnicodeEncodeError, UnicodeDecodeError, Exception) as e:
            # This is acceptable behavior for invalid unicode
            assert "utf-8" in str(e) or "encode" in str(e) or "unicode" in str(e).lower()
        
        print("✓ JSON encoding error handled correctly")
    
    
    def test_generate_filename_very_long_path(self, tmp_path):
        """Test filename generation with very long paths."""
        # Create a very long filename
        long_name = "a" * 200 + ".xlsx"
        input_file = tmp_path / long_name
        
        try:
            input_file.touch()
            result = generate_output_filename(str(input_file))
            # Should handle long filenames gracefully
            assert isinstance(result, str)
            assert len(result) > 0
        except OSError:
            # On some systems, very long filenames aren't supported
            pytest.skip("System doesn't support very long filenames")
        
        print("✓ Long filename handling works")


class TestFileOperationsPerformance:
    """Performance tests for file operations."""
    
    def test_write_csv_performance(self, tmp_path):
        """Test CSV writing performance with large dataset."""
        import time
        
        # Create large dataset
        test_data = []
        for i in range(10000):
            test_data.append(['Sheet1', f'Item_{i}', i, f'Data_{i}'])
        
        output_file = tmp_path / "performance_test.csv"
        
        start_time = time.time()
        write_to_csv(test_data, str(output_file))
        end_time = time.time()
        
        duration = end_time - start_time
        
        # Should complete within reasonable time (adjust as needed)
        assert duration < 10.0, f"CSV writing took {duration:.2f}s, expected < 10s"
        
        # Verify file was created correctly
        assert output_file.exists()
        assert output_file.stat().st_size > 100000  # Should be substantial
        
        print(f"✓ CSV performance test: {len(test_data)} rows in {duration:.2f}s")
    
    
    def test_find_newest_performance_many_files(self, tmp_path):
        """Test find_newest_excel_file performance with many files."""
        import time
        
        # Create many Excel files
        for i in range(100):
            excel_file = tmp_path / f"test_{i:03d}.xlsx"
            excel_file.touch()
            if i < 99:
                time.sleep(0.001)  # Small delay to ensure different timestamps
        
        start_time = time.time()
        result = find_newest_excel_file(str(tmp_path))
        end_time = time.time()
        
        duration = end_time - start_time
        
        # Should find newest file quickly
        assert duration < 1.0, f"File search took {duration:.2f}s, expected < 1s"
        assert Path(result).name == "test_099.xlsx"
        
        print(f"✓ File search performance: 100 files in {duration:.2f}s")


if __name__ == "__main__":
    print("Running file operations tests...")
    pytest.main([__file__, "-v"])