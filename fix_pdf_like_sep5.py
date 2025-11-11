"""
PDF Export Fix - Create proper workbook PDFs like September 5th output
Uses Microsoft Print to PDF driver for reliable conversion
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

def find_latest_output_dir():
    """Find the most recent output directory"""
    excel_files_dir = "Output_Record/Excel_Files"
    if os.path.exists(excel_files_dir):
        output_dirs = glob.glob(os.path.join(excel_files_dir, "output_*"))
        if output_dirs:
            # Filter out directories with locked files
            valid_dirs = []
            for dir_path in output_dirs:
                excel_files = glob.glob(os.path.join(dir_path, "*.xlsx"))
                excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
                if excel_files:
                    valid_dirs.append(dir_path)
            
            if valid_dirs:
                valid_dirs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return valid_dirs[0]
    return None

def convert_excel_to_proper_pdf(excel_file, output_dir):
    """Convert Excel workbook to PDF like September 5th method"""
    try:
        # Kill any existing Excel processes
        subprocess.run(['taskkill', '/f', '/im', 'EXCEL.EXE'], 
                      capture_output=True, check=False)
        time.sleep(2)
        
        print(f"Converting {os.path.basename(excel_file)}...")
        
        # Create Excel application
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open workbook
        workbook = excel.Workbooks.Open(os.path.abspath(excel_file))
        
        # Set up for PDF export (like September 5th)
        base_name = os.path.splitext(os.path.basename(excel_file))[0]
        pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
        
        # Export entire workbook as single PDF (this was the September 5th method)
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
            print(f"  ✓ Created: {base_name}.pdf ({file_size:.1f} KB)")
            return pdf_path
        else:
            print(f"  ✗ Failed to create PDF")
            return None
            
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return None

def main():
    print("=" * 80)
    print("PDF EXPORT FIX - September 5th Style Workbook PDFs")
    print("=" * 80)
    print("Creating proper workbook PDFs (not individual sheets)")
    print("=" * 80)
    
    # Find latest output directory
    output_dir = find_latest_output_dir()
    if not output_dir:
        print("No output directories found.")
        return
    
    print(f"Processing: {output_dir}")
    
    # Find Excel files
    excel_files = glob.glob(os.path.join(output_dir, "*.xlsx"))
    excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
    
    if not excel_files:
        print("No Excel files found.")
        return
        
    print(f"Found {len(excel_files)} Excel files:")
    for f in excel_files:
        print(f"  - {os.path.basename(f)}")
    
    # Create output directory
    pdf_dir = f"PDF_Output_Fixed_{datetime.now().strftime('%d%m%Y_%H%M')}"
    os.makedirs(pdf_dir, exist_ok=True)
    
    print(f"\nConverting to PDF...")
    print(f"Output directory: {pdf_dir}")
    
    # Convert each file
    converted_files = []
    for excel_file in excel_files:
        pdf_file = convert_excel_to_proper_pdf(excel_file, pdf_dir)
        if pdf_file:
            converted_files.append(pdf_file)
    
    print(f"\nConversion Results:")
    print(f"  Converted: {len(converted_files)}/{len(excel_files)} files")
    
    if converted_files:
        # Create ZIP archive
        zip_name = f"PDF_Export_Fixed_{datetime.now().strftime('%d%m%Y_%H%M')}.zip"
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for pdf_file in converted_files:
                arcname = os.path.basename(pdf_file)
                zipf.write(pdf_file, arcname)
                print(f"  Added to ZIP: {arcname}")
        
        zip_size = os.path.getsize(zip_name) / 1024 / 1024  # MB
        print(f"\n✓ Success!")
        print(f"  PDF Directory: {pdf_dir}")
        print(f"  ZIP Archive: {zip_name} ({zip_size:.1f} MB)")
        print(f"  Format: Proper workbook PDFs (like September 5th)")
        
        # Move to Output_Record
        if not os.path.exists("Output_Record/PDF_Files"):
            os.makedirs("Output_Record/PDF_Files")
        if not os.path.exists("Output_Record/PDF_Archives"):
            os.makedirs("Output_Record/PDF_Archives")
            
        import shutil
        shutil.move(pdf_dir, f"Output_Record/PDF_Files/{pdf_dir}")
        shutil.move(zip_name, f"Output_Record/PDF_Archives/{zip_name}")
        
        print(f"  Files moved to Output_Record structure")
    else:
        print("✗ No PDFs were created")

if __name__ == "__main__":
    main()