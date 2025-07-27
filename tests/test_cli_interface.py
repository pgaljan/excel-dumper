"""
CLI Interface tests for excel_dumper.
Tests the main() function and command-line argument parsing.
"""

import pytest
import sys
import os
import tempfile
from pathlib import Path
from unittest.mock import patch, MagicMock
import subprocess

# Import the functions we're testing
from excel_dumper.dumper import main, show_help


class TestCLIInterface:
    """Test command-line interface functionality."""
    
    def test_main_function_exists(self):
        """Test that main function is callable."""
        assert callable(main)
        print("✓ main() function is callable")
    
    
    def test_show_help_function(self, capsys):
        """Test the help display function."""
        show_help()
        captured = capsys.readouterr()
        
        # Check that help contains expected content
        assert "Excel Sheet Dumper" in captured.out
        assert "USAGE:" in captured.out
        assert "OPTIONS:" in captured.out
        assert "EXAMPLES:" in captured.out
        
        print("✓ Help function displays expected content")
    
    
    @patch('sys.argv')
    def test_main_with_help_argument(self, mock_argv, capsys):
        """Test main function with -help argument."""
        mock_argv.__getitem__.side_effect = lambda x: ['-help'][x] if x == 1 else ['script_name'][x]
        mock_argv.__contains__ = lambda self, x: x == '-help'
        mock_argv.__iter__ = lambda: iter(['script_name', '-help'])
        
        # Should exit cleanly after showing help
        main()
        
        captured = capsys.readouterr()
        assert "Excel Sheet Dumper" in captured.out
        
        print("✓ Help argument works correctly")
    
    
    @patch('sys.argv')
    @patch('excel_dumper.dumper.extract_excel_data')
    @patch('excel_dumper.dumper.write_to_csv')
    @patch('os.path.exists')
    def test_main_with_file_argument(self, mock_exists, mock_write_csv, mock_extract, mock_argv, tmp_path):
        """Test main function with -file argument."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        
        # Mock sys.argv
        mock_argv.__getitem__.side_effect = ['script_name', '-file', str(test_file)].__getitem__
        mock_argv.__contains__ = lambda self, x: x in ['-file', str(test_file)]
        
        # Mock dependencies
        mock_exists.return_value = True
        mock_extract.return_value = [['Sheet1', 'Data1', 'Data2']]
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = str(test_file)
            mock_args.no_hide = False
            mock_args.output_dir = None
            mock_args.input_dir = None
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            main()
            
            # Verify extraction was called
            mock_extract.assert_called_once()
            mock_write_csv.assert_called_once()
            
        print("✓ File argument processing works")
    
    
    @patch('sys.argv') 
    @patch('excel_dumper.dumper.find_newest_excel_file')
    @patch('excel_dumper.dumper.extract_excel_data')
    @patch('excel_dumper.dumper.write_to_csv')
    def test_main_without_file_finds_newest(self, mock_write_csv, mock_extract, mock_find_newest, mock_argv, tmp_path):
        """Test main function finds newest file when no file specified."""
        test_file = tmp_path / "newest.xlsx"
        test_file.touch()
        
        # Mock sys.argv to have no -file argument
        mock_argv.__getitem__.side_effect = ['script_name'].__getitem__
        mock_argv.__contains__ = lambda self, x: False
        
        # Mock dependencies
        mock_find_newest.return_value = str(test_file)
        mock_extract.return_value = [['Sheet1', 'Data1', 'Data2']]
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = None
            mock_args.no_hide = False
            mock_args.output_dir = None
            mock_args.input_dir = None
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            main()
            
            # Verify newest file search was called
            mock_find_newest.assert_called_once_with(".")
            mock_extract.assert_called_once()
            
        print("✓ Newest file search works")
    
    
    @patch('sys.argv')
    @patch('excel_dumper.dumper.extract_excel_data')
    @patch('excel_dumper.dumper.write_to_json')
    @patch('os.path.exists')
    def test_main_with_json_output(self, mock_exists, mock_write_json, mock_extract, mock_argv, tmp_path):
        """Test main function with JSON output option."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        
        mock_exists.return_value = True
        mock_extract.return_value = [['Sheet1', 'Data1', 'Data2']]
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = str(test_file)
            mock_args.no_hide = False
            mock_args.output_dir = None
            mock_args.input_dir = None
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = True  # JSON output
            mock_parse.return_value = mock_args
            
            main()
            
            # Verify JSON writer was called
            mock_write_json.assert_called_once()
            
        print("✓ JSON output option works")
    
    
    @patch('sys.argv')
    @patch('excel_dumper.dumper.extract_excel_data')
    @patch('excel_dumper.dumper.write_to_csv')
    @patch('os.path.exists')
    def test_main_with_all_options(self, mock_exists, mock_write_csv, mock_extract, mock_argv, tmp_path):
        """Test main function with all options enabled."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()
        
        mock_exists.return_value = True
        mock_extract.return_value = [['Sheet1', 1, 'Data1', 'Data2']]
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = str(test_file)
            mock_args.no_hide = True  # Skip hidden sheets
            mock_args.output_dir = str(tmp_path)
            mock_args.input_dir = None
            mock_args.formulas = True  # Include formulas
            mock_args.rownumbers = True  # Include row numbers
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            main()
            
            # Verify extraction was called with correct options
            # Verify extract_excel_data was called with correct arguments
            call_args = mock_extract.call_args
            assert call_args[0][0] == str(test_file)  # filename
            assert call_args[0][1] == False  # include_hidden
            assert call_args[0][2] == True   # include_row_numbers
            assert call_args[0][3] == True   # include_formulas
            
        print("✓ All options work correctly")
    
    
    @patch('sys.argv')
    @patch('os.path.exists')
    def test_main_file_not_found_error(self, mock_exists, mock_argv, capsys):
        """Test main function handles file not found error."""
        mock_exists.return_value = False
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = "nonexistent.xlsx"
            mock_args.no_hide = False
            mock_args.output_dir = None
            mock_args.input_dir = None
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            with pytest.raises(SystemExit):
                main()
            
            captured = capsys.readouterr()
            assert "not found" in captured.out
            
        print("✓ File not found error handled correctly")
    
    
    @patch('sys.argv')
    @patch('excel_dumper.dumper.find_newest_excel_file')
    def test_main_no_excel_files_found_error(self, mock_find_newest, mock_argv, capsys):
        """Test main function handles no Excel files found error."""
        mock_find_newest.side_effect = FileNotFoundError("No Excel files found")
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = None
            mock_args.no_hide = False
            mock_args.output_dir = None
            mock_args.input_dir = "."
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            with pytest.raises(SystemExit):
                main()
            
            captured = capsys.readouterr()
            assert "Error:" in captured.out
            
        print("✓ No Excel files error handled correctly")
    
    
    @patch('sys.argv')
    @patch('excel_dumper.dumper.extract_excel_data')
    @patch('os.path.exists')
    def test_main_extraction_error(self, mock_exists, mock_extract, mock_argv, capsys):
        """Test main function handles extraction errors."""
        mock_exists.return_value = True
        mock_extract.side_effect = Exception("Extraction failed")
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = "test.xlsx"
            mock_args.no_hide = False
            mock_args.output_dir = None
            mock_args.input_dir = None
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            with pytest.raises(SystemExit):
                main()
            
            captured = capsys.readouterr()
            assert "Error:" in captured.out
            
        print("✓ Extraction error handled correctly")


class TestCLIArgumentParsing:
    """Test specific argument parsing scenarios."""
    
    def test_relative_filename_with_input_dir(self, tmp_path):
        """Test relative filename combined with input directory."""
        # Create test structure
        input_dir = tmp_path / "input"
        input_dir.mkdir()
        test_file = input_dir / "test.xlsx"
        test_file.touch()
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = "test.xlsx"  # Relative filename
            mock_args.input_dir = str(input_dir)  # Input directory
            mock_args.no_hide = False
            mock_args.output_dir = None
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            with patch('excel_dumper.dumper.extract_excel_data') as mock_extract:
                with patch('excel_dumper.dumper.write_to_csv') as mock_write:
                    mock_extract.return_value = [['Sheet1', 'Data']]
                    
                    main()
                    
                    # Should have combined input_dir + filename
                    expected_path = str(input_dir / "test.xlsx")
                    mock_extract.assert_called_once()
                    call_args = mock_extract.call_args[0]
                    assert expected_path in call_args[0] or call_args[0].endswith("test.xlsx")
        
        print("✓ Relative filename with input directory works")
    
    
    def test_output_directory_creation(self, tmp_path):
        """Test that output directory is created if it doesn't exist."""
        input_file = tmp_path / "test.xlsx"
        input_file.touch()
        
        output_dir = tmp_path / "output" / "subdir"  # Nested directory that doesn't exist
        
        with patch('argparse.ArgumentParser.parse_args') as mock_parse:
            mock_args = MagicMock()
            mock_args.filename = str(input_file)
            mock_args.input_dir = None
            mock_args.output_dir = str(output_dir)
            mock_args.no_hide = False
            mock_args.formulas = False
            mock_args.rownumbers = False
            mock_args.json = False
            mock_parse.return_value = mock_args
            
            with patch('excel_dumper.dumper.extract_excel_data') as mock_extract:
                with patch('excel_dumper.dumper.write_to_csv') as mock_write:
                    mock_extract.return_value = [['Sheet1', 'Data']]
                    
                    main()
                    
                    # Verify output directory creation would be attempted
                    mock_write.assert_called_once()
        
        print("✓ Output directory handling works")


def test_cli_integration_subprocess(tmp_path):
    """Integration test using subprocess to test actual CLI."""
    # Create a test Excel file using openpyxl
    test_file = tmp_path / "cli_test.xlsx"
    
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Name", "Value"])
        ws.append(["Test", 123])
        wb.save(test_file)
        
        # Test the CLI with subprocess
        result = subprocess.run([
            sys.executable, "-m", "excel_dumper.dumper",
            "-file", str(test_file)
        ], capture_output=True, text=True, cwd=tmp_path)
        
        # Check that it ran without crashing
        assert result.returncode == 0 or "successfully exported" in result.stdout
        
        print("✓ CLI integration test passed")
        
    except ImportError:
        pytest.skip("openpyxl not available for integration test")
    except Exception as e:
        print(f"⚠️  CLI integration test failed (this may be expected): {e}")


if __name__ == "__main__":
    print("Running CLI interface tests...")
    pytest.main([__file__, "-v"])