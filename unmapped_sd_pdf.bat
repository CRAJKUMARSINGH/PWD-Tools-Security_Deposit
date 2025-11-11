@echo off
setlocal EnableExtensions EnableDelayedExpansion
echo ================================================================================
echo UNMAPPED SECURITY DEPOSIT REFUND WITH PDF GENERATION
echo ================================================================================
echo PWD Electric Division - Udaipur
echo Developer: RAJKUMAR SINGH CHAUHAN
echo ================================================================================

REM Check if Python is available
python -c "import sys; print(sys.version)" >NUL 2>&1
if errorlevel 1 (
    echo Error: Python not found in PATH.
    echo Please ensure Python is installed and added to PATH.
    pause
    exit /b 2
)

REM Check if required input file exists
if not exist "deductions_unmapped.xlsx" (
    echo Error: deductions_unmapped.xlsx not found.
    echo Please ensure the unmapped deductions file is in the current directory.
    pause
    exit /b 1
)

echo Step 1: Running Unmapped Security Deposit Processor...
echo ================================================================================
python sd_deduction_for_unmapped.py
if errorlevel 1 (
    echo Error: Unmapped SD processing failed with error level %errorlevel%
    pause
    exit /b %errorlevel%
)

echo.
echo Step 2: Converting Excel files to PDF...
echo ================================================================================
python simple_pdf_export.py
if errorlevel 1 (
    echo Error: PDF export failed with error level %errorlevel%
    pause
    exit /b %errorlevel%
)

echo.
echo ================================================================================
echo SUCCESS: Unmapped Security Deposit Refund with PDF Generation Completed!
echo ================================================================================
echo Check the following outputs:
echo - Excel files: SD_Deductions_Summary_*.xlsx
echo - PDF files: PDF_Output_*\ directory  
echo - ZIP archive: PDF_Export_*.zip
echo ================================================================================
pause
endlocal