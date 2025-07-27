#!/usr/bin/env python3
"""
Direct test runner that includes all test files.
"""

import subprocess
import sys
from pathlib import Path

def run_all_tests_direct():
    """Run all available test files directly."""
    
    # List all test files
    test_files = [
        "test_excel_extraction.py",
        "test_example_usage.py", 
        "test_cli_interface.py",
        "test_file_operations.py",
        "test_coverage_boost.py"
    ]
    
    # Find existing test files
    existing_files = []
    for test_file in test_files:
        if Path(test_file).exists():
            existing_files.append(test_file)
            print(f"✓ Found: {test_file}")
        else:
            print(f"✗ Missing: {test_file}")
    
    if not existing_files:
        print("❌ No test files found!")
        return False
    
    print(f"\n🚀 Running {len(existing_files)} test files with coverage...")
    
    # Run pytest with coverage
    cmd = [
        sys.executable, "-m", "pytest",
        "--cov=excel_dumper",
        "--cov-report=term-missing",
        "--cov-report=html",
        "--cov-report=xml", 
        "-v",
        "--tb=short"
    ] + existing_files
    
    print("Executing:", " ".join(cmd[:8]) + "... [test files]")
    print("=" * 60)
    
    try:
        result = subprocess.run(cmd, capture_output=False, text=True)
        
        if result.returncode == 0:
            print("\n🎉 ALL TESTS PASSED!")
            print("\n📊 Coverage Analysis:")
            print("- Coverage report: htmlcov/index.html")
            print("- XML report: coverage.xml")
            return True
        else:
            print(f"\n❌ Tests failed with return code: {result.returncode}")
            return False
            
    except Exception as e:
        print(f"❌ Error running tests: {e}")
        return False


if __name__ == "__main__":
    run_all_tests_direct()
