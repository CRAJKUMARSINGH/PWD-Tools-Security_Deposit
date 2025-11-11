@echo off
setlocal EnableExtensions EnableDelayedExpansion
echo ================================================================================
echo PDF EXPORT FIX - Creating Proper Workbook PDFs like September 5th
echo ================================================================================
echo PWD Electric Division - Udaipur
echo Developer: RAJKUMAR SINGH CHAUHAN
echo ================================================================================

REM Find the latest output directory with Excel files
set "LATEST_DIR="
for /f "delims=" %%D in ('dir "Output_Record\Excel_Files\output_*" /b /ad /o-d 2^>nul') do (
    if not defined LATEST_DIR set "LATEST_DIR=%%D"
)

if "%LATEST_DIR%"=="" (
    echo No output directories found in Output_Record\Excel_Files\
    pause
    exit /b 1
)

echo Latest output directory: %LATEST_DIR%
set "SOURCE_DIR=Output_Record\Excel_Files\%LATEST_DIR%"
set "PDF_DIR=PDF_Output_Fixed_%LATEST_DIR%"

echo Source: %SOURCE_DIR%
echo PDF Output: %PDF_DIR%

REM Create PDF output directory
if not exist "%PDF_DIR%" mkdir "%PDF_DIR%"

echo.
echo Converting Excel files to proper workbook PDFs...
echo ================================================================================

REM Use Microsoft Print to PDF (available on Windows 10+)
for %%F in ("%SOURCE_DIR%\*.xlsx") do (
    if not "%%~nF"=="" (
        echo Converting: %%~nxF
        REM Use PowerShell to open Excel and print to PDF
        powershell -Command "& {
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $workbook = $excel.Workbooks.Open('%%~fF')
                $pdfPath = '%CD%\%PDF_DIR%\%%~nF.pdf'
                $workbook.ExportAsFixedFormat(0, $pdfPath, 0)
                $workbook.Close()
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                Write-Host '  Success: %%~nF.pdf created'
            } catch {
                Write-Host '  Error: Failed to convert %%~nF - ' $_.Exception.Message
            }
        }"
    )
)

echo.
echo ================================================================================
echo Creating ZIP archive...
echo ================================================================================

REM Create ZIP of all PDFs
set "ZIP_NAME=PDF_Export_Fixed_%LATEST_DIR%.zip"
powershell -Command "Compress-Archive -Path '%PDF_DIR%\*.pdf' -DestinationPath '%ZIP_NAME%' -Force"

if exist "%ZIP_NAME%" (
    echo Success: Created %ZIP_NAME%
    for %%F in ("%ZIP_NAME%") do echo Archive size: %%~zF bytes
) else (
    echo Error: Failed to create ZIP archive
)

echo.
echo ================================================================================
echo PDF Export Fix Completed!
echo ================================================================================
echo Check the following outputs:
echo - PDF files: %PDF_DIR%\
echo - ZIP archive: %ZIP_NAME%
echo - Format: Proper workbook PDFs (like September 5th)
echo ================================================================================
pause
endlocal