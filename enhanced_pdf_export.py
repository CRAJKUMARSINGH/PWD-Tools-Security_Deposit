"""
Enhanced PDF Export - Create proper workbook PDFs like September 5th output
Converts entire Excel workbooks to single PDF files (not individual sheets)
PWD Electric Division - Udaipur
Developer: RAJKUMAR SINGH CHAUHAN
"""

import os
import sys
import glob
from pathlib import Path
import subprocess
import zipfile
from datetime import datetime

def convert_excel_to_pdf_libreoffice(excel_file, output_dir):
    """
    Convert Excel workbook to PDF using LibreOffice (maintains workbook format)
    This creates one PDF per workbook (like September 5th output)
    """
    try:
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"Converting {excel_file} using LibreOffice...")
        
        # Get workbook name without extension
        base_name = Path(excel_file).stem
        pdf_filename = f"{base_name}.pdf"
        
        # LibreOffice command to convert Excel to PDF
        cmd = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            os.path.abspath(excel_file)
        ]
        
        # Try LibreOffice conversion
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        
        if result.returncode == 0:
            pdf_path = os.path.join(output_dir, pdf_filename)
            if os.path.exists(pdf_path):
                print(f"  ✓ Converted workbook to {pdf_filename}")
                return [pdf_path]
            else:
                print(f"  ✗ PDF file not created: {pdf_filename}")
                return []
        else:
            print(f"  ✗ LibreOffice conversion failed: {result.stderr}")
            return []
            
    except subprocess.TimeoutExpired:
        print(f"  ✗ LibreOffice conversion timed out for {excel_file}")
        return []
    except FileNotFoundError:
        print(f"  ✗ LibreOffice not found. Please install LibreOffice.")
        return []
    except Exception as e:
        print(f"  ✗ Error converting {excel_file}: {e}")
        return []

def convert_excel_to_pdf_win32com(excel_file, output_dir):
    """
    Convert Excel workbook to PDF using Excel COM automation (Windows only)
    This creates one PDF per workbook (preserves September 5th format)
    """
    try:
        import win32com.client as win32
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"Converting {excel_file} using Excel COM...")
        
        # Kill any existing Excel processes
        try:
            subprocess.run(['taskkill', '/f', '/im', 'EXCEL.EXE'], 
                         capture_output=True, check=False)
        except:
            pass
        
        # Initialize Excel application
        excel_app = win32.DispatchEx('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.EnableEvents = False
        
        # Open workbook
        workbook = excel_app.Workbooks.Open(os.path.abspath(excel_file))
        
        # Get workbook name without extension
        base_name = Path(excel_file).stem
        pdf_filename = f"{base_name}.pdf"
        pdf_path = os.path.join(output_dir, pdf_filename)
        
        # Export entire workbook to PDF
        workbook.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=pdf_path,
            Quality=0,  # xlQualityStandard
            From=1,
            To=workbook.Worksheets.Count,
            OpenAfterPublish=False
        )
        
        # Close workbook and Excel
        workbook.Close(False)
        excel_app.Quit()
        
        # Clean up COM objects
        try:
            del workbook
            del excel_app
        except:
            pass
        
        if os.path.exists(pdf_path):
            print(f"  ✓ Converted workbook to {pdf_filename}")
            return [pdf_path]
        else:
            print(f"  ✗ PDF file not created: {pdf_filename}")
            return []
            
    except ImportError:
        print("  ✗ win32com not available. Installing pywin32...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pywin32"], check=True)
        return convert_excel_to_pdf_win32com(excel_file, output_dir)
    except Exception as e:
        print(f"  ✗ Error converting {excel_file}: {e}")
        return []

def find_latest_output_dir():
    """Find the most recent output_* directory in Output_Record/Excel_Files"""
    # Check Output_Record/Excel_Files first
    excel_files_dir = "Output_Record/Excel_Files"
    if os.path.exists(excel_files_dir):
        output_dirs = glob.glob(os.path.join(excel_files_dir, "output_*"))
        if output_dirs:
            # Sort by modification time, newest first
            output_dirs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return output_dirs[0]
    
    # Fallback to current directory
    output_dirs = glob.glob("output_*")
    if not output_dirs:
        return None
    
    # Sort by modification time, newest first
    output_dirs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return output_dirs[0]

def create_zip_archive(pdf_dir, zip_name=None):
    """Create a ZIP archive of all PDF files"""
    if not zip_name:
        timestamp = datetime.now().strftime("%d%m%Y_%H%M")
        zip_name = f"PDF_Export_Workbook_{timestamp}.zip"
    
    pdf_files = glob.glob(os.path.join(pdf_dir, "*.pdf"))
    
    if not pdf_files:
        print("No PDF files found to archive.")
        return None
    
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for pdf_file in pdf_files:
            arcname = os.path.basename(pdf_file)
            zipf.write(pdf_file, arcname)
            print(f"  Added {arcname} to archive")
    
    print(f"✓ Created ZIP archive: {zip_name}")
    print(f"  Archive size: {os.path.getsize(zip_name) / 1024 / 1024:.1f} MB")
    return zip_name

def main():
    print("=" * 80)
    print("ENHANCED PDF EXPORT UTILITY - WORKBOOK FORMAT")
    print("=" * 80)
    print("This creates PDFs like the September 5th output (one PDF per workbook)")
    print("=" * 80)
    
    # Find the latest output directory
    output_dir = find_latest_output_dir()
    if not output_dir:
        print("No output directories found. Please run the security deposit generator first.")
        return
    
    print(f"Using latest output directory: {output_dir}")
    
    # Find all Excel files in the output directory
    excel_files = glob.glob(os.path.join(output_dir, "*.xlsx"))
    
    if not excel_files:
        print(f"No Excel files found in {output_dir}")
        return
    
    print(f"Found {len(excel_files)} Excel file(s) to convert:")
    for excel_file in excel_files:
        print(f"  - {os.path.basename(excel_file)}")
    
    # Create PDF output directory
    pdf_dir = f"PDF_Output_Workbook_{os.path.basename(output_dir)}"
    
    print(f"\nConverting to PDF (workbook format)...")
    print(f"PDF output directory: {pdf_dir}")
    
    all_converted_files = []
    
    # Try LibreOffice first, then fall back to Excel COM
    conversion_methods = [
        ("LibreOffice", convert_excel_to_pdf_libreoffice),
        ("Excel COM", convert_excel_to_pdf_win32com)
    ]
    
    for method_name, convert_func in conversion_methods:
        print(f"\nTrying {method_name} conversion...")
        
        for excel_file in excel_files:
            converted_files = convert_func(excel_file, pdf_dir)
            all_converted_files.extend(converted_files)
        
        if all_converted_files:
            print(f"✓ {method_name} conversion successful!")
            break
        else:
            print(f"✗ {method_name} conversion failed, trying next method...")
    
    print(f"\nConversion Summary:")
    print(f"  Total PDF files created: {len(all_converted_files)}")
    print(f"  Format: Workbook PDFs (like September 5th output)")
    
    if all_converted_files:
        # Create ZIP archive
        print(f"\nCreating ZIP archive...")
        zip_file = create_zip_archive(pdf_dir)
        
        if zip_file:
            print(f"\n✓ PDF Export completed successfully!")
            print(f"  PDF Directory: {pdf_dir}")
            print(f"  ZIP Archive: {zip_file}")
            print(f"  Format: Proper workbook PDFs (like September 5th)")
        
    else:
        print("✗ No PDF files were created.")
        print("Please ensure LibreOffice is installed or Excel is available.")

if __name__ == "__main__":
    main()