# Blank Security Refund Generator

## Overview
This folder contains the **Enhanced Blank Security Refund Generator** that creates blank security deposit refund forms for PWD Electric Division - Udaipur contractors.

## Key Improvements Over Original
✅ **No Hardcoded Paths** - Automatically searches for input files in multiple locations  
✅ **Self-Contained** - All required files in one folder  
✅ **Better Error Handling** - Clear error messages and validation  
✅ **Professional Output** - Clean, organized blank forms ready for manual filling  
✅ **Configurable** - Easy to modify and extend  

## Files in This Folder
- `enhanced_blank_generator.py` - Main generator script (improved version)
- `run_blank_generator.bat` - Batch file to run the generator
- `work_order_master.xlsx` - Input data file (copy of main file)
- `README.md` - This documentation

## How to Use

### Method 1: Double-click the batch file
```
run_blank_generator.bat
```

### Method 2: Command line
```bash
cd Blank_Generator
python enhanced_blank_generator.py
```

## Input Requirements
- `work_order_master.xlsx` (automatically detected in current folder or parent directories)
- The script will search for the file in multiple locations automatically

## Output
- Creates a timestamped folder: `BLANK_SD_SHEETS_DD-MM-YYYY_HH-MM/`
- Generates multiple Excel workbooks (25 sheets per workbook)
- Each sheet is a blank security deposit refund form
- Ready for manual data entry
- Print-ready A4 portrait format

## Features
- **Automatic File Detection**: No need to modify paths
- **Batch Processing**: Processes all work orders efficiently
- **Professional Formatting**: Proper borders, fonts, and layout
- **Print Ready**: A4 portrait format, fits on one page
- **Error Handling**: Clear error messages and validation
- **Timestamped Output**: Organized output with date/time stamps

## Solved Issues
❌ **Original Issue**: Hardcoded path `C:\Users\Rajkumar\Security_Deposit\work_order_master.xlsx`  
✅ **Solution**: Auto-detection of input files in multiple standard locations

❌ **Original Issue**: Required manual path modification for different environments  
✅ **Solution**: Self-contained folder with relative path resolution

❌ **Original Issue**: Poor error handling when files not found  
✅ **Solution**: Comprehensive validation and user-friendly error messages

## Developer
**RAJKUMAR SINGH CHAUHAN**  
Email: crajkumarsingh@hotmail.com  
PWD Electric Division - Udaipur

## Version History
- **v2.0** (Current) - Enhanced version with auto-path detection and improved error handling
- **v1.0** - Original version with hardcoded paths (deprecated)

---
*This generator is part of the PWD-Tools-Security_Deposit project for automating security deposit refund processing.*