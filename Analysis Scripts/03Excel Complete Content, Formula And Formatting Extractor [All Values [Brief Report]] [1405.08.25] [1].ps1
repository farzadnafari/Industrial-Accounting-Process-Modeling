# ====================================================================================================
# PROJECT       : Industrial Cost Accounting – Process Modeling And Notation
# SCRIPT        : Excel Compact Data And Formula Extractor
# CATEGORY      : Excel Data Extraction / Compact Workbook Documentation
# PLATFORM      : Windows PowerShell | Microsoft Excel | COM Automation
# ====================================================================================================
# PURPOSE:
# Extracts non-empty cell contents, formulas, and calculated results from an Excel workbook.
# Produces a compact, structured TXT report designed for quick inspection and further analysis.
#
# EXTRACTION FEATURES:
# [1] CELL ADDRESSES       : Identifies the address of every extracted cell.
# [2] CELL VALUES          : Captures text, numbers, and other available cell values.
# [3] FORMULAS             : Extracts formulas assigned to individual cells.
# [4] FORMULA RESULTS      : Includes the currently available calculated result of each formula.
# [5] WORKSHEET COVERAGE   : Automatically processes all worksheets in the selected workbook.
# [6] COMPACT FORMAT       : Groups extracted cells by row to reduce report length.
# [7] EMPTY CELL FILTER    : Ignores cells without values or formulas.
#
# EXECUTION WORKFLOW:
# [1] INPUT                : Requests the full path of the source Excel workbook.
# [2] VALIDATION           : Verifies that the specified file exists.
# [3] INITIALIZATION       : Starts Microsoft Excel in background mode.
# [4] READ-ONLY ACCESS     : Opens the workbook without intentionally modifying its contents.
# [5] DATA INSPECTION      : Scans the used range of each worksheet.
# [6] DATA FORMATTING      : Converts extracted content into a compact text representation.
# [7] REPORT GENERATION    : Writes the results into a UTF-8 encoded text file.
# [8] RESOURCE CLEANUP     : Closes Excel and releases COM objects.
#
# REPORT STRUCTURE:
# - Source workbook path and extraction date.
# - Individual worksheet sections.
# - Cell addresses and corresponding values.
# - Formula expressions followed by their available calculated results.
# - Compact, row-oriented records separated by vertical bars.
#
# OUTPUT SPECIFICATION:
# FORMAT                 : Plain Text (.txt)
# ENCODING               : UTF-8 with BOM
# LOCATION               : Same directory as the source Excel workbook
# FILE NAMING             : [WorkbookName]_Compact.txt
# REPORT TYPE             : Compact Report
#
# REQUIREMENTS:
# - Windows operating system.
# - Microsoft Excel desktop application installed.
# - PowerShell with Excel COM Automation support.
# - Read access to the source workbook and write access to its directory.
#
# IMPORTANT NOTES:
# - The source workbook is opened in read-only mode.
# - Formula results reflect the values currently available through Excel.
# - Line breaks are replaced with spaces, and vertical bars inside cell content are replaced with slashes.
# - The report is intended for quick inspection, documentation, and preliminary data analysis.
# - This utility does not create, modify, or overwrite worksheets in the source workbook.
#
# ====================================================================================================
#                               END OF SCRIPT DOCUMENTATION
# ====================================================================================================

$ErrorActionPreference = "Stop"

Write-Host "`nEXCEL COMPACT EXTRACTOR" -ForegroundColor Cyan

$ExcelPath = Read-Host "Enter Excel file path"

if (-not (Test-Path -LiteralPath $ExcelPath -PathType Leaf)) {
    Write-Host "File not found!" -ForegroundColor Red
    Read-Host "Press ENTER to exit"
    exit
}

$ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path
$Folder = Split-Path -Parent $ExcelPath
$BaseName = [IO.Path]::GetFileNameWithoutExtension($ExcelPath)
$OutputPath = Join-Path $Folder ($BaseName + "_Compact.txt")

$Excel = $null
$Workbook = $null
$Writer = $null

function Format-CellValue {
    param($Value)

    if ($null -eq $Value) { return "" }

    return ([string]$Value).
        Replace("`r", " ").
        Replace("`n", " ").
        Replace("|", "/")
}

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

    $Workbook = $Excel.Workbooks.Open($ExcelPath, 0, $true)

    $Writer.WriteLine("EXCEL COMPACT DATA REPORT")
    $Writer.WriteLine("FILE: $ExcelPath")
    $Writer.WriteLine("DATE: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    $Writer.WriteLine("========================================")

    foreach ($Sheet in $Workbook.Worksheets) {

        Write-Host "Processing: $($Sheet.Name)" -ForegroundColor Green

        $Writer.WriteLine("")
        $Writer.WriteLine("[SHEET: $($Sheet.Name)]")

        $Range = $Sheet.UsedRange
        $StartRow = $Range.Row
        $StartCol = $Range.Column
        $EndRow = $StartRow + $Range.Rows.Count - 1
        $EndCol = $StartCol + $Range.Columns.Count - 1

        for ($r = $StartRow; $r -le $EndRow; $r++) {

            $Items = [System.Collections.Generic.List[string]]::new()

            for ($c = $StartCol; $c -le $EndCol; $c++) {

                $Cell = $Sheet.Cells.Item($r, $c)

                $Value = $Cell.Value2
                $HasFormula = $Cell.HasFormula

                if ($null -eq $Value -and -not $HasFormula) {
                    continue
                }

                $Address = $Cell.Address($false, $false)
                $ValueText = Format-CellValue $Value

                if ($HasFormula) {
                    $Formula = Format-CellValue $Cell.Formula
                    $Items.Add("$Address=$Formula => $ValueText")
                }
                else {
                    $Items.Add("$Address=$ValueText")
                }
            }

            if ($Items.Count -gt 0) {
                $Writer.WriteLine(($Items -join " | "))
            }
        }
    }

    $Writer.WriteLine("")
    $Writer.WriteLine("========== END OF REPORT ==========")

    Write-Host "`nExtraction completed." -ForegroundColor Green
    Write-Host "Output: $OutputPath" -ForegroundColor Cyan
}
catch {
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
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
    }

    if ($null -ne $Excel) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Read-Host "`nPress ENTER to exit"