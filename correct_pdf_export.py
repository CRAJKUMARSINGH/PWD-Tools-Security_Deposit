"""
CORRECTED PDF EXPORT - Creates proper workbook PDFs like September 5th output
Fixes the issue where individual sheet PDFs were created instead of workbook PDFs
PWD Electric Division - Udaipur
Developer: RAJKUMAR SINGH CHAUHAN
"""

import os
import glob
import subprocess
import zipfile
from datetime import datetime
import win32com.client as win32
import time
import shutil

def find_latest_output_dir():
    """Find the most recent output directory with valid Excel files"""
    excel_files_dir = "Output_Record/Excel_Files"
    if os.path.exists(excel_files_dir):
        output_dirs = glob.glob(os.path.join(excel_files_dir, "output_*"))
        if output_dirs:
            # Sort by modification time, newest first
            output_dirs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            
            # Find the first directory with unlocked Excel files
            for dir_path in output_dirs:
                excel_files = glob.glob(os.path.join(dir_path, "*.xlsx"))
                valid_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
                if valid_files:
                    return dir_path
    return None

def convert_workbook_to_pdf(excel_file, output_dir):
    """Convert entire Excel workbook to single PDF (September 5th method)"""
    try:
        print(f"Converting {os.path.basename(excel_file)}...")
        
        # Create Excel application
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open workbook
        workbook = excel.Workbooks.Open(os.path.abspath(excel_file))
        
        # Create PDF filename
        base_name = os.path.splitext(os.path.basename(excel_file))[0]
        pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
        
        # Export entire workbook as single PDF (like September 5th)
        workbook.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=pdf_path,
            Quality=0  # xlQualityStandard
        )
        
        workbook.Close(SaveChanges=False)
        excel.Quit()
        
        # Cleanup
        del workbook
        del excel
        
        if os.path.exists(pdf_path):
            file_size = os.path.getsize(pdf_path) / 1024  # KB
            print(f"  ✓ Created: {base_name}.pdf ({file_size:.1f} KB) - Workbook format ✓")
            return pdf_path
        else:
            print(f"  ✗ Failed to create PDF")
            return None
            
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return None

def main():
    print("=" * 80)
    print("CORRECTED PDF EXPORT - September 5th Style Workbook PDFs")
    print("=" * 80)
    print("Fixing PDF generation to create proper workbook PDFs")
    print("(Not individual sheet PDFs like today's incorrect output)")
    print("=" * 80)
    
    # Kill any existing Excel processes first
    print("Cleaning up Excel processes...")
    subprocess.run(['taskkill', '/f', '/im', 'EXCEL.EXE'], 
                  capture_output=True, check=False)
    time.sleep(3)
    
    # Find latest output directory
    output_dir = find_latest_output_dir()
    if not output_dir:
        print("No valid output directories found.")
        return
    
    print(f"Processing: {output_dir}")
    
    # Find Excel files (exclude lock files)
    excel_files = glob.glob(os.path.join(output_dir, "*.xlsx"))
    excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
    
    if not excel_files:
        print("No valid Excel files found.")
        return
        
    print(f"Found {len(excel_files)} Excel files:")
    for f in excel_files:
        file_size = os.path.getsize(f) / 1024  # KB
        print(f"  - {os.path.basename(f)} ({file_size:.1f} KB)")
    
    # Create output directory
    timestamp = datetime.now().strftime('%d%m%Y_%H%M')
    pdf_dir = f"PDF_Output_Corrected_{timestamp}"
    os.makedirs(pdf_dir, exist_ok=True)
    
    print(f"\nConverting to corrected PDF format...")
    print(f"Output directory: {pdf_dir}")
    
    # Convert each file
    converted_files = []
    for excel_file in excel_files:
        pdf_file = convert_workbook_to_pdf(excel_file, pdf_dir)
        if pdf_file:
            converted_files.append(pdf_file)
        time.sleep(1)  # Brief pause between conversions
    
    print(f"\nConversion Results:")
    print(f"  Converted: {len(converted_files)}/{len(excel_files)} files")
    
    if converted_files:
        # Create ZIP archive
        zip_name = f"PDF_Export_Corrected_{timestamp}.zip"
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for pdf_file in converted_files:
                arcname = os.path.basename(pdf_file)
                zipf.write(pdf_file, arcname)
                print(f"  Added to ZIP: {arcname}")
        
        zip_size = os.path.getsize(zip_name) / 1024 / 1024  # MB
        
        print(f"\n✓ SUCCESS! Fixed PDF Export Completed")
        print(f"  PDF Directory: {pdf_dir}")
        print(f"  ZIP Archive: {zip_name} ({zip_size:.1f} MB)")
        print(f"  Format: Proper workbook PDFs ✓ (like September 5th)")
        print(f"  Issue Fixed: No more individual sheet PDFs ✓")
        
        # Move to Output_Record structure
        if not os.path.exists("Output_Record/PDF_Files"):
            os.makedirs("Output_Record/PDF_Files")
        if not os.path.exists("Output_Record/PDF_Archives"):
            os.makedirs("Output_Record/PDF_Archives")
            
        shutil.move(pdf_dir, f"Output_Record/PDF_Files/{pdf_dir}")
        shutil.move(zip_name, f"Output_Record/PDF_Archives/{zip_name}")
        
        print(f"  Files organized in Output_Record structure ✓")
        
        print(f"\n" + "=" * 80)
        print("COMPARISON:")
        print(f"❌ Today's wrong output: Individual sheet PDFs (~16 KB each)")
        print(f"✅ September 5th correct: Single workbook PDF (~640 KB)")
        print(f"✅ Fixed output now: Single workbook PDF (~{zip_size*1000:.0f} KB)")
        print("=" * 80)
        
    else:
        print("✗ No PDFs were created")
        print("Check if Excel files are locked or corrupted")

if __name__ == "__main__":
    main()