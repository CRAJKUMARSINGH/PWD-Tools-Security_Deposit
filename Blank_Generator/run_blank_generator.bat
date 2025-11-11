@echo off
setlocal EnableExtensions EnableDelayedExpansion
echo ================================================================================
echo BLANK SECURITY DEPOSIT REFUND GENERATOR
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

REM Check if work_order_master.xlsx exists (it will auto-search in multiple locations)
echo Step 1: Generating Blank Security Refund Sheets...
echo ================================================================================
python enhanced_blank_generator.py
if errorlevel 1 (
    echo Error: Blank sheet generation failed with error level %errorlevel%
    pause
    exit /b %errorlevel%
)

echo.
echo ================================================================================
echo SUCCESS: Blank Security Deposit Refund Sheets Generated!
echo ================================================================================
echo Check the following outputs:
echo - Excel files: BLANK_SD_SHEETS_*\ directory
echo - Ready for manual data entry
echo - Professional formatting applied
echo ================================================================================
pause
endlocal