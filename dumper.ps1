#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Excel Sheet Dumper Script - PowerShell Version
    Extracts all non-null rows from Excel worksheets and saves to CSV format.

.DESCRIPTION
    This script replicates the functionality of the Python dumper.py script.
    It processes Excel files (.xlsx, .xls, .xlsm, .xlsb) and extracts data
    from all worksheets into a single CSV file with metadata.

.PARAMETER File
    Specify Excel file to process (default: newest Excel file in input directory)

.PARAMETER InputDir
    Input directory to search for Excel files (default: current directory)

.PARAMETER OutputDir
    Output directory for CSV file (default: current directory)

.PARAMETER NoHide
    Skip hidden worksheets (default: include all worksheets)

.PARAMETER RowNumbers
    Include Excel row numbers in output (default: exclude row numbers)

.PARAMETER Formulas
    Show formulas instead of calculated values (.xlsx/.xlsm only)

.PARAMETER Help
    Show detailed help message

.EXAMPLE
    .\dumper.ps1
    Process newest Excel file from current directory

.EXAMPLE
    .\dumper.ps1 -File "data.xlsx"
    Process specific file

.EXAMPLE
    .\dumper.ps1 -RowNumbers
    Include Excel row numbers in output

.EXAMPLE
    .\dumper.ps1 -Formulas
    Show formulas instead of values

.EXAMPLE
    .\dumper.ps1 -InputDir "./source" -OutputDir "./exports"
    Source and output directories

.EXAMPLE
    .\dumper.ps1 -File "data.xlsx" -OutputDir "./exports" -NoHide -RowNumbers -Formulas
    All options combined
#>

param(
    [string]$File,
    [string]$InputDir = ".",
    [string]$OutputDir,
    [switch]$NoHide,
    [switch]$RowNumbers,
    [switch]$Formulas,
    [switch]$Help
)

# Show detailed help if requested
if ($Help) {
    $helpText = @"
Excel Sheet Dumper - PowerShell Version
Extract data from Excel worksheets to CSV

USAGE:
    .\dumper.ps1 [OPTIONS]

OPTIONS:
    -File FILE          Specify Excel file to process (default: newest Excel file in input directory)
    -InputDir DIR       Input directory to search for Excel files (default: current directory)
    -OutputDir DIR      Output directory for CSV file (default: current directory)
    -NoHide             Skip hidden worksheets (default: include all worksheets)
    -RowNumbers         Include Excel row numbers in output (default: exclude row numbers)
    -Formulas           Show formulas instead of calculated values (.xlsx/.xlsm only)
    -Help               Show this help message

EXAMPLES:
    .\dumper.ps1                                    # Process newest Excel file from current directory
    .\dumper.ps1 -File "data.xlsx"                  # Process specific file
    .\dumper.ps1 -RowNumbers                        # Include Excel row numbers in output
    .\dumper.ps1 -Formulas                          # Show formulas instead of values
    .\dumper.ps1 -InputDir "./source"               # Process newest file from ./source directory
    .\dumper.ps1 -InputDir "./source" -OutputDir "./exports"  # Source and output directories
    .\dumper.ps1 -InputDir "/data" -File "report.xlsx" -RowNumbers     # Specific file with row numbers
    .\dumper.ps1 -File "data.xlsx" -OutputDir "./exports" -NoHide -RowNumbers -Formulas  # All options combined

OUTPUT:
    Creates a CSV file named "dumper_[original_filename]_[timestamp].csv" with:
    - Timestamp is the last modified time of the originating Excel file
    - Timestamp format: ISO 8601 with colons replaced by hyphens
    - If file exists, appends incremental number in parentheses
    - First column: Worksheet name
    - Second column: Excel row number (if -RowNumbers option used)
    - Remaining columns: Original data from worksheets
    - Only non-empty rows are included
    - Formulas are prefixed with 'FORMULA: =' to prevent circular references

NOTE: The -Formulas option only works with .xlsx and .xlsm files. For .xls and .xlsb files, 
calculated values will be shown regardless of this setting.

POWERSHELL REQUIREMENTS:
    - ImportExcel module (Install-Module ImportExcel -Force)
    - PowerShell 5.1 or PowerShell Core 6+

SUPPORTED EXCEL FORMATS:
    - .xlsx (Excel 2007+)
    - .xls  (Excel 97-2003)
    - .xlsm (Excel Macro-Enabled)
    - .xlsb (Excel Binary)
"@
    Write-Host $helpText
    return
}

# Check if ImportExcel module is available
try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Error "Error: ImportExcel module not found."
    Write-Host "Please install with: Install-Module ImportExcel -Force"
    Write-Host "You may need to run PowerShell as Administrator for the first installation."
    exit 1
}

function Find-NewestExcelFile {
    param([string]$SearchDir = ".")
    
    $excelExtensions = @("*.xlsx", "*.xls", "*.xlsm", "*.xlsb")
    $excelFiles = @()
    
    foreach ($extension in $excelExtensions) {
        $files = Get-ChildItem -Path $SearchDir -Filter $extension -File -ErrorAction SilentlyContinue
        $excelFiles += $files
    }
    
    if ($excelFiles.Count -eq 0) {
        throw "No Excel files found in directory: $SearchDir"
    }
    
    # Get the newest file based on LastWriteTime
    $newestFile = $excelFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    return $newestFile.FullName
}

function Test-NonNullData {
    param([array]$Row)
    
    foreach ($cell in $Row) {
        if ($null -ne $cell -and $cell.ToString().Trim() -ne '') {
            return $true
        }
    }
    return $false
}

function Extract-ExcelData {
    param(
        [string]$FileName,
        [bool]$IncludeHidden = $true,
        [bool]$IncludeRowNumbers = $false,
        [bool]$IncludeFormulas = $false
    )
    
    $extractedData = @()
    
    try {
        # Check for unsupported file types with formulas option
        $fileExt = [System.IO.Path]::GetExtension($FileName).ToLower()
        if ($IncludeFormulas -and $fileExt -notin @('.xlsx', '.xlsm')) {
            Write-Warning "Formulas option only works with .xlsx and .xlsm files."
            Write-Warning "File '$FileName' will be processed with calculated values instead of formulas."
            $IncludeFormulas = $false
        }
        
        # Get all worksheet names
        $workbook = Open-ExcelPackage -Path $FileName
        $worksheetNames = $workbook.Workbook.Worksheets | ForEach-Object { $_.Name }
        
        foreach ($sheetName in $worksheetNames) {
            try {
                $worksheet = $workbook.Workbook.Worksheets[$sheetName]
                
                # Check if sheet is hidden
                if (-not $IncludeHidden -and $worksheet.Hidden -ne [OfficeOpenXml.eWorkSheetHidden]::Visible) {
                    Write-Host "Skipping hidden sheet: $sheetName"
                    continue
                }
                
                # Get the used range
                if ($worksheet.Dimension -eq $null) {
                    continue  # Skip empty sheets
                }
                
                $startRow = $worksheet.Dimension.Start.Row
                $endRow = $worksheet.Dimension.End.Row
                $startCol = $worksheet.Dimension.Start.Column
                $endCol = $worksheet.Dimension.End.Column
                
                # Process each row
                for ($rowNum = $startRow; $rowNum -le $endRow; $rowNum++) {
                    $rowData = @()
                    $hasData = $false
                    
                    for ($colNum = $startCol; $colNum -le $endCol; $colNum++) {
                        $cell = $worksheet.Cells[$rowNum, $colNum]
                        
                        if ($null -ne $cell.Value) {
                            if ($IncludeFormulas -and -not [string]::IsNullOrEmpty($cell.Formula)) {
                                # Prefix with "FORMULA: " to prevent CSV interpretation
                                $rowData += "FORMULA: =$($cell.Formula)"
                            } else {
                                $rowData += $cell.Value
                            }
                            $hasData = $true
                        } else {
                            $rowData += $null
                        }
                    }
                    
                    if ($hasData -and (Test-NonNullData -Row $rowData)) {
                        # Build the output row
                        if ($IncludeRowNumbers) {
                            $rowWithMetadata = @($sheetName, $rowNum) + $rowData
                        } else {
                            $rowWithMetadata = @($sheetName) + $rowData
                        }
                        
                        $extractedData += ,@($rowWithMetadata)
                    }
                }
                
            } catch {
                Write-Warning "Could not process sheet '$sheetName': $($_.Exception.Message)"
                continue
            }
        }
        
        $workbook.Dispose()
        
    } catch {
        throw "Error reading Excel file '$FileName': $($_.Exception.Message)"
    }
    
    return $extractedData
}

function Write-ToCsv {
    param(
        [array]$Data,
        [string]$OutputFileName,
        [bool]$IncludeRowNumbers = $false
    )
    
    try {
        # Determine maximum number of columns in the data
        $maxCols = 0
        foreach ($row in $Data) {
            if ($row.Count -gt $maxCols) {
                $maxCols = $row.Count
            }
        }
        
        # Create header
        if ($Data.Count -gt 0) {
            if ($IncludeRowNumbers) {
                # Structure: [worksheet_name, row_number, ...data_columns...]
                $dataCols = $maxCols - 2
                if ($dataCols -gt 0) {
                    $header = @('Worksheet', 'Row_Number') + (1..$dataCols | ForEach-Object { "Column_$_" })
                } else {
                    $header = @('Worksheet', 'Row_Number')
                }
            } else {
                # Structure: [worksheet_name, ...data_columns...]
                $dataCols = $maxCols - 1
                if ($dataCols -gt 0) {
                    $header = @('Worksheet') + (1..$dataCols | ForEach-Object { "Column_$_" })
                } else {
                    $header = @('Worksheet')
                }
            }
            
            # Write header to CSV
            $header -join ',' | Out-File -FilePath $OutputFileName -Encoding UTF8
        }
        
        # Write data rows
        foreach ($row in $Data) {
            # Escape and quote fields that contain commas, quotes, or newlines
            $escapedRow = @()
            foreach ($field in $row) {
                if ($null -eq $field) {
                    $escapedRow += ''
                } else {
                    $fieldStr = $field.ToString()
                    if ($fieldStr -match '[",\r\n]') {
                        $fieldStr = $fieldStr -replace '"', '""'
                        $escapedRow += "`"$fieldStr`""
                    } else {
                        $escapedRow += $fieldStr
                    }
                }
            }
            $escapedRow -join ',' | Out-File -FilePath $OutputFileName -Append -Encoding UTF8
        }
        
        Write-Host "Data successfully exported to: $OutputFileName"
        Write-Host "Total rows exported: $($Data.Count)"
        
    } catch {
        throw "Error writing to CSV file '$OutputFileName': $($_.Exception.Message)"
    }
}

function New-OutputFileName {
    param(
        [string]$InputFileName,
        [string]$OutputDir
    )
    
    $inputFile = Get-Item $InputFileName
    $baseName = $inputFile.BaseName
    
    # Get file modification time with timezone
    $modTime = $inputFile.LastWriteTime
    
    # Format timestamp as ISO 8601 with colons replaced by hyphens
    # Use custom format to avoid colon in timezone offset
    $timestamp = $modTime.ToString("yyyy-MM-ddTHH-mm-sszzz") -replace ":", ""
    
    $baseFileName = "dumperps_$($baseName)_$timestamp"
    
    # Determine the directory path
    if ($OutputDir) {
        if (-not (Test-Path $OutputDir)) {
            New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
        }
        $basePath = Join-Path $OutputDir $baseFileName
    } else {
        $basePath = $baseFileName
    }
    
    # Check if file exists and find an available filename
    $counter = 0
    do {
        if ($counter -eq 0) {
            $finalFileName = "$basePath.csv"
        } else {
            $finalFileName = "$basePath($counter).csv"
        }
        $counter++
    } while (Test-Path $finalFileName)
    
    return $finalFileName
}

# Main execution
try {
    # Determine input file
    if ($File) {
        # If filename is provided, check if it's absolute or relative
        if (-not [System.IO.Path]::IsPathRooted($File) -and $InputDir -ne ".") {
            # If it's relative and InputDir is specified, join them
            $inputFile = Join-Path $InputDir $File
        } else {
            $inputFile = $File
        }
        
        if (-not (Test-Path $inputFile)) {
            Write-Error "File '$inputFile' not found."
            exit 1
        }
    } else {
        try {
            $inputFile = Find-NewestExcelFile -SearchDir $InputDir
            $fileName = Split-Path $inputFile -Leaf
            $fileDir = Split-Path $inputFile -Parent
            Write-Host "Processing newest Excel file: $fileName"
            Write-Host "From directory: $fileDir"
        } catch {
            Write-Error $_.Exception.Message
            exit 1
        }
    }
    
    # Extract data
    $includeHidden = -not $NoHide
    Write-Host "Extracting data from: $inputFile"
    Write-Host "Including hidden sheets: $includeHidden"
    Write-Host "Including row numbers: $RowNumbers"
    Write-Host "Including formulas: $Formulas"
    
    $extractedData = Extract-ExcelData -FileName $inputFile -IncludeHidden $includeHidden -IncludeRowNumbers $RowNumbers -IncludeFormulas $Formulas
    
    if ($extractedData.Count -eq 0) {
        Write-Host "No data found to export."
        return
    }
    
    # Generate output filename and write CSV
    $outputFile = New-OutputFileName -InputFileName $inputFile -OutputDir $OutputDir
    Write-ToCsv -Data $extractedData -OutputFileName $outputFile -IncludeRowNumbers $RowNumbers
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    exit 1
}