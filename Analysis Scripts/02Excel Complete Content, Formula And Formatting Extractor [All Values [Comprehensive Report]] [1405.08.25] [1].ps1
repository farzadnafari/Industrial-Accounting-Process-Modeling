# ====================================================================================================
# PROJECT       : Industrial Cost Accounting – Process Modeling And Notation
# SCRIPT        : Excel Cell Values And Formulas Extractor
# CATEGORY      : Excel Workbook Inspection / Data Structure Documentation
# PLATFORM      : Windows PowerShell | Microsoft Excel | COM Automation
# ====================================================================================================
# PURPOSE:
# This utility extracts cell contents, formulas, and calculated formula results from an Excel workbook.
# It creates a structured, human-readable TXT report to support workbook inspection and documentation.
#
# EXTRACTION OPTIONS:
# [1] CELL ADDRESSES       : Records the address of each non-empty cell in every worksheet.
# [2] STATIC VALUES        : Extracts text, numbers, dates, and other stored cell values.
# [3] FORMULAS             : Captures the formulas defined in individual worksheet cells.
# [4] CALCULATED RESULTS   : Records the currently available results returned by Excel formulas.
# [5] WORKSHEETS           : Automatically processes all worksheets in the selected workbook.
# [6] EMPTY CELLS          : Skips cells that contain neither a value nor a formula.
#
# EXECUTION WORKFLOW:
# [1] INPUT                : Requests the full path of the source Excel workbook.
# [2] VALIDATION           : Checks whether the specified file exists.
# [3] EXCEL INITIALIZATION : Starts Microsoft Excel in background mode.
# [4] WORKBOOK ACCESS      : Opens the source workbook in read-only mode.
# [5] DATA EXTRACTION      : Inspects the used range of each worksheet, cell by cell.
# [6] REPORT GENERATION    : Writes extracted information into a structured UTF-8 text file.
# [7] RESOURCE CLEANUP     : Closes Excel objects and releases COM resources.
#
# OUTPUT SPECIFICATION:
# FORMAT                 : Plain Text (.txt)
# ENCODING               : UTF-8 with BOM
# LOCATION               : Same directory as the source Excel workbook
# FILE NAMING             : [WorkbookName]_Values_And_Formulas.txt
# REPORT CONTENTS         : Worksheet names, cell addresses, values, formulas, and calculated results.
#
# REQUIREMENTS:
# - Windows operating system.
# - Microsoft Excel desktop application installed.
# - Permission to access the source workbook and create the output report.
# - PowerShell with access to Excel COM Automation.
#
# IMPORTANT NOTES:
# - The original Excel workbook is opened in read-only mode and is not intentionally modified.
# - Formula results represent the values currently available through Excel's calculation engine.
# - External links, macros, and data connections are not intentionally refreshed by this utility.
# - The extraction is intended for technical inspection, auditing, documentation, and analysis.
# - This script generates a text-based report; it does not create or modify Excel worksheets.
#
# ====================================================================================================
#                               END OF SCRIPT DOCUMENTATION
# ====================================================================================================

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "EXCEL VALUES AND FORMULAS EXTRACTOR" -ForegroundColor Cyan
Write-Host "==================================="

$ExcelPath = Read-Host "Enter the full Excel file path"

if (-not (Test-Path -LiteralPath $ExcelPath -PathType Leaf)) {
    Write-Host "ERROR: Excel file not found." -ForegroundColor Red
    Read-Host "Press ENTER to exit"
    exit
}

$ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path

$Folder = Split-Path -Parent $ExcelPath
$FileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath)

$OutputPath = Join-Path $Folder ($FileName + "_Values_And_Formulas.txt")

$Excel = $null
$Workbook = $null
$Writer = $null

try {

    $Writer = New-Object System.IO.StreamWriter(
        $OutputPath,
        $false,
        [System.Text.UTF8Encoding]::new($true)
    )

    $Excel = New-Object -ComObject Excel.Application

    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.EnableEvents = $false
    $Excel.AskToUpdateLinks = $false

    Write-Host ""
    Write-Host "Opening Excel workbook..." -ForegroundColor Yellow

    # Open workbook as read-only
    $Workbook = $Excel.Workbooks.Open($ExcelPath, 0, $true)

    $Writer.WriteLine("EXCEL CONTENTS AND FORMULAS REPORT")
    $Writer.WriteLine("Source: $ExcelPath")
    $Writer.WriteLine("Date: $(Get-Date)")
    $Writer.WriteLine("=" * 80)

    # Process every worksheet
    foreach ($Sheet in $Workbook.Worksheets) {

        Write-Host "Processing: $($Sheet.Name)" -ForegroundColor Green

        $Writer.WriteLine("")
        $Writer.WriteLine("=" * 80)
        $Writer.WriteLine("WORKSHEET: $($Sheet.Name)")
        $Writer.WriteLine("=" * 80)

        $UsedRange = $Sheet.UsedRange

        $StartRow = $UsedRange.Row
        $StartCol = $UsedRange.Column

        $EndRow = $StartRow + $UsedRange.Rows.Count - 1
        $EndCol = $StartCol + $UsedRange.Columns.Count - 1

        for ($Row = $StartRow; $Row -le $EndRow; $Row++) {

            for ($Col = $StartCol; $Col -le $EndCol; $Col++) {

                $Cell = $Sheet.Cells.Item($Row, $Col)

                $Value = $Cell.Value2
                $Formula = $Cell.Formula
                $HasFormula = $Cell.HasFormula

                # Skip completely empty cells
                if ($null -eq $Value -and -not $HasFormula) {
                    continue
                }

                $Address = $Cell.Address($false, $false)

                $Writer.WriteLine("")
                $Writer.WriteLine("----------------------------------------")
                $Writer.WriteLine("CELL: $Address")
                $Writer.WriteLine("----------------------------------------")

                if ($HasFormula) {

                    $Writer.WriteLine("TYPE: FORMULA")
                    $Writer.WriteLine("FORMULA: $Formula")
                    $Writer.WriteLine("CALCULATED RESULT: $Value")

                }
                else {

                    $Writer.WriteLine("TYPE: VALUE")
                    $Writer.WriteLine("VALUE: $Value")

                }
            }
        }

        $Writer.WriteLine("")
        $Writer.WriteLine("END OF WORKSHEET: $($Sheet.Name)")
    }

    $Writer.WriteLine("")
    $Writer.WriteLine("=" * 80)
    $Writer.WriteLine("EXTRACTION COMPLETED")
    $Writer.WriteLine("=" * 80)

    Write-Host ""
    Write-Host "Extraction completed successfully." -ForegroundColor Green
    Write-Host "Output: $OutputPath" -ForegroundColor Cyan

}
catch {

    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red

}
finally {

    if ($null -ne $Workbook) {
        $Workbook.Close($false)
    }

    if ($null -ne $Excel) {
        $Excel.Quit()
    }

    if ($null -ne $Writer) {
        $Writer.Flush()
        $Writer.Close()
    }

    if ($null -ne $Workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    }

    if ($null -ne $Excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Read-Host "Press ENTER to exit"