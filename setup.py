#!/usr/bin/env python3
"""
Setup script for excel-dumper package.
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="excel-dumper",
    version="1.0.0",
    author="pgaljan",
    author_email="galjan@gmail.com",
    description="Cross-platform Excel ETL preprocessor for data pipeline ingestion and auditing",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers", 
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pandas>=1.5.0",
        "openpyxl>=3.0.0", 
        "xlrd>=2.0.0",
    ],
    entry_points={
        "console_scripts": [
            "excel-dumper=excel_dumper.dumper:main",
            "dumper=excel_dumper.dumper:main",
        ],
    },
)
