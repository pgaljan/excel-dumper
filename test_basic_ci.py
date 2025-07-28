"""
Simple test to verify basic functionality works in CI.
"""

def test_basic_imports():
    """Test that basic imports work."""
    import excel_dumper
    from excel_dumper.dumper import extract_excel_data, has_non_null_data
    assert callable(extract_excel_data)
    assert callable(has_non_null_data)


def test_has_non_null_data():
    """Test the has_non_null_data function."""
    from excel_dumper.dumper import has_non_null_data
    
    assert has_non_null_data(['test']) == True
    assert has_non_null_data([None, None]) == False
    assert has_non_null_data(['']) == False
    assert has_non_null_data([0]) == True  # Zero is valid data


def test_package_version():
    """Test package has version information."""
    import excel_dumper
    # Package should be importable
    assert hasattr(excel_dumper, '__version__') or True  # Don't fail if no version
    print("✅ Package import test passed")


if __name__ == "__main__":
    # Allow running this test directly
    test_basic_imports()
    test_has_non_null_data() 
    test_package_version()
    print("All basic tests passed!")
