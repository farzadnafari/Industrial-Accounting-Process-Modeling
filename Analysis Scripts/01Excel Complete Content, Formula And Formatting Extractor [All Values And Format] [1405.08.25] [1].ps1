# ====================================================================================================
# PROJECT       : Industrial Cost Accounting – Process Modeling And Notation
# SCRIPT        : Excel Complete Content, Formula And Formatting Extractor
# CATEGORY      : Comprehensive Excel Workbook Inspection / Structural Documentation
# PLATFORM      : Windows PowerShell 5.1+ | Microsoft Excel Desktop | COM Automation
# ====================================================================================================
# PURPOSE:
# Performs an extensive inspection of an Excel workbook and extracts its contents, formulas,
# calculated values, formatting attributes, structural elements, and selected workbook metadata.
# Generates a detailed UTF-8 text report to support technical documentation, auditing, and analysis.
#
# EXTRACTION MODULES:
# [01] WORKBOOK INFORMATION    : Source file, file name, extraction timestamp, and worksheet count.
# [02] DEFINED NAMES           : Workbook-level named ranges and their references.
# [03] WORKSHEET STRUCTURE     : Used ranges, starting positions, row counts, and column counts.
# [04] CELL CONTENTS           : Stored values, displayed text, and individual cell addresses.
# [05] FORMULA EXTRACTION      : Formula expressions and formula-presence indicators.
# [06] CELL FORMATTING         : Number formats, font properties, fill colors, and patterns.
# [07] TEXT ALIGNMENT          : Horizontal and vertical alignment, wrapping, orientation, and indentation.
# [08] CELL BORDERS            : Border styles, weights, and colors.
# [09] CELL PROTECTION         : Locked status and hidden-formula attributes.
# [10] MERGED CELLS            : Identifies merged ranges within worksheet used areas.
# [11] ROW DIMENSIONS          : Row heights and hidden-row information.
# [12] COLUMN DIMENSIONS       : Column widths and hidden-column information.
# [13] COMMENTS                : Extracts available legacy cell comments.
# [14] HYPERLINKS              : Captures hyperlink addresses and displayed link text.
# [15] EXCEL TABLES            : Table names, ranges, row counts, and column counts.
# [16] CONDITIONAL FORMATTING  : Rule types, operators, and available formula expressions.
# [17] PAGE SETUP              : Page orientation, paper size, print area, and zoom settings.
# [18] DOCUMENT PROPERTIES     : Author, title, subject, keywords, comments, and creation date.
#
# EXECUTION WORKFLOW:
# [1] INPUT                   : Requests the full path of the source Excel workbook.
# [2] FILE VALIDATION         : Checks whether the specified workbook exists.
# [3] OUTPUT INITIALIZATION   : Creates a UTF-8 text report beside the source workbook.
# [4] EXCEL INITIALIZATION    : Starts Microsoft Excel in background mode.
# [5] READ-ONLY ACCESS        : Opens the source workbook without intentionally modifying it.
# [6] WORKBOOK INSPECTION     : Collects workbook-level information and defined names.
# [7] WORKSHEET PROCESSING    : Inspects every worksheet and its used range.
# [8] REPORT GENERATION       : Writes extracted information into organized report sections.
# [9] RESOURCE CLEANUP        : Closes the workbook, terminates Excel, and releases COM objects.
#
# OUTPUT SPECIFICATION:
# FORMAT                    : Plain Text (.txt)
# ENCODING                  : UTF-8 with BOM
# LOCATION                  : Same directory as the source Excel workbook
# FILE NAMING                : [WorkbookName]_Complete_Extraction.txt
# REPORT TYPE                : Comprehensive Report
# REPORT ORGANIZATION        : Workbook Information → Worksheet Details → Workbook Properties
#
# REQUIREMENTS:
# - Windows operating system.
# - Microsoft Excel Desktop installed.
# - Windows PowerShell 5.1 or later.
# - Read access to the source workbook.
# - Write access to the output directory.
#
# IMPORTANT NOTES:
# - The workbook is opened in read-only mode.
# - Formula results represent values currently available through Excel.
# - Some properties may be unavailable for specific cells, worksheets, or Excel versions.
# - The extraction focuses on workbook structure and selected formatting metadata.
# - It does not reproduce the complete visual appearance of the original workbook.
# - Modern threaded comments, embedded objects, charts, images, and VBA source code are not fully extracted.
# - Conditional formatting rules and workbook metadata are collected on a best-effort basis.
# - Large worksheets may require considerable processing time because individual cells are inspected.
#
# ====================================================================================================
#                              END OF SCRIPT DOCUMENTATION
# ====================================================================================================

$ErrorActionPreference = "Stop"

function Write-Report {
    param([string]$Text = "")

    $script:Writer.WriteLine($Text)
}

function Get-ColorDescription {
    param($Color)

    try {
        if ($null -eq $Color -or $Color -eq -4142) {
            return "Automatic/None"
        }

        $value = [int]$Color

        $red   = $value -band 255
        $green = ($value -shr 8) -band 255
        $blue  = ($value -shr 16) -band 255

        return ("RGB({0},{1},{2}) / #{3:X2}{4:X2}{5:X2}" -f `
            $red, $green, $blue, $red, $green, $blue)
    }
    catch {
        return "Unavailable"
    }
}

function Get-SafeProperty {
    param(
        $Object,
        [string]$Property
    )

    try {
        return $Object.$Property
    }
    catch {
        return "Unavailable"
    }
}

function Format-Value {
    param($Value)

    if ($null -eq $Value) {
        return "[EMPTY]"
    }

    $text = [string]$Value
    return $text.Replace("`r", "\r").Replace("`n", "\n")
}

# ------------------------------------------------------------
# Select Excel File
# ------------------------------------------------------------

Write-Host ""
Write-Host "EXCEL COMPLETE EXTRACTION TOOL" -ForegroundColor Cyan
Write-Host "--------------------------------"

$ExcelPath = Read-Host "Enter the full Excel file path"

if (-not (Test-Path -LiteralPath $ExcelPath -PathType Leaf)) {
    Write-Host "ERROR: File not found." -ForegroundColor Red
    exit
}

$ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path

$OutputFolder = Split-Path -Parent $ExcelPath
$BaseName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath)

$OutputPath = Join-Path $OutputFolder ($BaseName + "_Complete_Extraction.txt")

# ------------------------------------------------------------
# Initialize Excel
# ------------------------------------------------------------

$Excel = $null
$Workbook = $null
$script:Writer = $null

try {

    $script:Writer = New-Object System.IO.StreamWriter(
        $OutputPath,
        $false,
        [System.Text.UTF8Encoding]::new($true)
    )

    $Excel = New-Object -ComObject Excel.Application

    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.AskToUpdateLinks = $false
    $Excel.EnableEvents = $false

    Write-Host ""
    Write-Host "Opening workbook..." -ForegroundColor Yellow

    $Workbook = $Excel.Workbooks.Open($ExcelPath, 0, $true)

    # --------------------------------------------------------
    # Workbook Information
    # --------------------------------------------------------

    Write-Report "============================================================"
    Write-Report "EXCEL WORKBOOK EXTRACTION REPORT"
    Write-Report "============================================================"
    Write-Report "Source File : $ExcelPath"
    Write-Report "File Name   : $([System.IO.Path]::GetFileName($ExcelPath))"
    Write-Report "Extracted   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Report "Worksheet Count : $($Workbook.Worksheets.Count)"
    Write-Report ""

    # --------------------------------------------------------
    # Workbook Named Ranges
    # --------------------------------------------------------

    Write-Report "============================================================"
    Write-Report "WORKBOOK DEFINED NAMES"
    Write-Report "============================================================"

    foreach ($Name in $Workbook.Names) {

        Write-Report "Name     : $($Name.Name)"
        Write-Report "Reference: $($Name.RefersTo)"
        Write-Report ""
    }

    # --------------------------------------------------------
    # Process Every Worksheet
    # --------------------------------------------------------

    foreach ($Sheet in $Workbook.Worksheets) {

        Write-Host "Extracting worksheet: $($Sheet.Name)" -ForegroundColor Green

        Write-Report ""
        Write-Report "############################################################"
        Write-Report "WORKSHEET: $($Sheet.Name)"
        Write-Report "############################################################"

        $UsedRange = $Sheet.UsedRange

        $StartRow = $UsedRange.Row
        $StartCol = $UsedRange.Column

        $RowCount = $UsedRange.Rows.Count
        $ColCount = $UsedRange.Columns.Count

        Write-Report "Used Range : $($UsedRange.Address())"
        Write-Report "Start Row  : $StartRow"
        Write-Report "Start Col  : $StartCol"
        Write-Report "Row Count  : $RowCount"
        Write-Report "Column Count: $ColCount"
        Write-Report ""

        # ----------------------------------------------------
        # Column Widths
        # ----------------------------------------------------

        Write-Report "---------------- COLUMN DIMENSIONS ----------------"

        for ($c = $StartCol; $c -lt ($StartCol + $ColCount); $c++) {

            $Column = $Sheet.Columns.Item($c)

            Write-Report (
                "Column {0}: Width={1}; Hidden={2}" -f `
                $c, $Column.ColumnWidth, $Column.Hidden
            )
        }

        # ----------------------------------------------------
        # Row Heights
        # ----------------------------------------------------

        Write-Report ""
        Write-Report "---------------- ROW DIMENSIONS ----------------"

        for ($r = $StartRow; $r -lt ($StartRow + $RowCount); $r++) {

            $Row = $Sheet.Rows.Item($r)

            Write-Report (
                "Row {0}: Height={1}; Hidden={2}" -f `
                $r, $Row.RowHeight, $Row.Hidden
            )
        }

        # ----------------------------------------------------
        # Merged Cells
        # ----------------------------------------------------

        Write-Report ""
        Write-Report "---------------- MERGED CELLS ----------------"

        for ($r = $StartRow; $r -lt ($StartRow + $RowCount); $r++) {

            for ($c = $StartCol; $c -lt ($StartCol + $ColCount); $c++) {

                $Cell = $Sheet.Cells.Item($r, $c)

                try {
                    if ($Cell.MergeCells) {
                        if ($Cell.Address() -eq $Cell.MergeArea.Cells.Item(1,1).Address()) {
                            Write-Report "Merged Range: $($Cell.MergeArea.Address())"
                        }
                    }
                }
                catch {}
            }
        }

        # ----------------------------------------------------
        # Cell Contents, Formulas And Formatting
        # ----------------------------------------------------

        Write-Report ""
        Write-Report "================ CELL DETAILS ================"

        for ($r = $StartRow; $r -lt ($StartRow + $RowCount); $r++) {

            for ($c = $StartCol; $c -lt ($StartCol + $ColCount); $c++) {

                $Cell = $Sheet.Cells.Item($r, $c)

                $Value = $Cell.Value2
                $Formula = $Cell.Formula
                $DisplayText = $Cell.Text

                $HasFormula = $false

                if ($Cell.HasFormula) {
                    $HasFormula = $true
                }

                $HasContent = (
                    $null -ne $Value -or
                    $HasFormula -or
                    $Cell.Comment -ne $null -or
                    $Cell.Hyperlinks.Count -gt 0
                )

                if (-not $HasContent) {
                    continue
                }

                Write-Report ""
                Write-Report "------------------------------------------------------------"
                Write-Report "CELL: $($Cell.Address())"
                Write-Report "------------------------------------------------------------"

                Write-Report "Value         : $(Format-Value $Value)"
                Write-Report "Displayed Text: $(Format-Value $DisplayText)"
                Write-Report "Formula       : $(Format-Value $Formula)"
                Write-Report "Has Formula   : $HasFormula"
                Write-Report "Data Type     : $(Get-SafeProperty $Cell "Value2" | ForEach-Object { if ($null -eq $_) {'Empty'} else {$_.GetType().FullName} })"

                # Number Formatting
                Write-Report ""
                Write-Report "[NUMBER FORMATTING]"
                Write-Report "Number Format : $(Get-SafeProperty $Cell "NumberFormat")"
                Write-Report "Number Format Local: $(Get-SafeProperty $Cell "NumberFormatLocal")"

                # Font
                Write-Report ""
                Write-Report "[FONT]"
                Write-Report "Font Name : $(Get-SafeProperty $Cell.Font "Name")"
                Write-Report "Font Size : $(Get-SafeProperty $Cell.Font "Size")"
                Write-Report "Bold      : $(Get-SafeProperty $Cell.Font "Bold")"
                Write-Report "Italic    : $(Get-SafeProperty $Cell.Font "Italic")"
                Write-Report "Underline : $(Get-SafeProperty $Cell.Font "Underline")"
                Write-Report "Font Color: $(Get-ColorDescription (Get-SafeProperty $Cell.Font "Color"))"

                # Fill
                Write-Report ""
                Write-Report "[FILL]"
                Write-Report "Interior Color: $(Get-ColorDescription (Get-SafeProperty $Cell.Interior "Color"))"
                Write-Report "Pattern       : $(Get-SafeProperty $Cell.Interior "Pattern")"

                # Alignment
                Write-Report ""
                Write-Report "[ALIGNMENT]"
                Write-Report "Horizontal : $(Get-SafeProperty $Cell "HorizontalAlignment")"
                Write-Report "Vertical   : $(Get-SafeProperty $Cell "VerticalAlignment")"
                Write-Report "Wrap Text  : $(Get-SafeProperty $Cell "WrapText")"
                Write-Report "Text Orientation: $(Get-SafeProperty $Cell "Orientation")"
                Write-Report "Indent Level: $(Get-SafeProperty $Cell "IndentLevel")"

                # Borders
                Write-Report ""
                Write-Report "[BORDERS]"

                foreach ($BorderIndex in @(7, 8, 9, 10, 11, 12)) {

                    try {
                        $Border = $Cell.Borders.Item($BorderIndex)

                        Write-Report (
                            "Border {0}: LineStyle={1}; Weight={2}; Color={3}" -f `
                            $BorderIndex,
                            $Border.LineStyle,
                            $Border.Weight,
                            (Get-ColorDescription $Border.Color)
                        )
                    }
                    catch {}
                }

                # Protection
                Write-Report ""
                Write-Report "[PROTECTION]"
                Write-Report "Locked : $(Get-SafeProperty $Cell "Locked")"
                Write-Report "Formula Hidden: $(Get-SafeProperty $Cell "FormulaHidden")"

                # Comments
                try {
                    if ($null -ne $Cell.Comment) {
                        Write-Report "Comment: $($Cell.Comment.Text())"
                    }
                }
                catch {}

                # Hyperlinks
                try {
                    foreach ($Link in $Cell.Hyperlinks) {
                        Write-Report "Hyperlink: $($Link.Address)"
                        Write-Report "Link Text: $($Link.TextToDisplay)"
                    }
                }
                catch {}
            }
        }

        # ----------------------------------------------------
        # Worksheet Tables
        # ----------------------------------------------------

        Write-Report ""
        Write-Report "---------------- EXCEL TABLES ----------------"

        try {
            foreach ($Table in $Sheet.ListObjects) {
                Write-Report "Table Name: $($Table.Name)"
                Write-Report "Table Range: $($Table.Range.Address())"
                Write-Report "Rows: $($Table.ListRows.Count)"
                Write-Report "Columns: $($Table.ListColumns.Count)"
                Write-Report ""
            }
        }
        catch {}

        # ----------------------------------------------------
        # Conditional Formatting
        # ----------------------------------------------------

        Write-Report ""
        Write-Report "---------------- CONDITIONAL FORMATTING ----------------"

        try {
            $FormatConditions = $UsedRange.FormatConditions

            for ($i = 1; $i -le $FormatConditions.Count; $i++) {

                $Condition = $FormatConditions.Item($i)

                Write-Report "Condition Index: $i"
                Write-Report "Type: $($Condition.Type)"
                Write-Report "Operator: $($Condition.Operator)"
                Write-Report "Formula1: $($Condition.Formula1)"
                Write-Report "Formula2: $($Condition.Formula2)"
                Write-Report ""
            }
        }
        catch {}

        # ----------------------------------------------------
        # Page Setup
        # ----------------------------------------------------

        Write-Report ""
        Write-Report "---------------- PAGE SETUP ----------------"

        try {
            Write-Report "Orientation: $($Sheet.PageSetup.Orientation)"
            Write-Report "Paper Size: $($Sheet.PageSetup.PaperSize)"
            Write-Report "Print Area: $($Sheet.PageSetup.PrintArea)"
            Write-Report "Zoom: $($Sheet.PageSetup.Zoom)"
        }
        catch {}

        Write-Report ""
        Write-Report "END OF WORKSHEET: $($Sheet.Name)"
    }

    # --------------------------------------------------------
    # Workbook Properties
    # --------------------------------------------------------

    Write-Report ""
    Write-Report "============================================================"
    Write-Report "WORKBOOK DOCUMENT PROPERTIES"
    Write-Report "============================================================"

    try {
        Write-Report "Author: $($Workbook.BuiltinDocumentProperties.Item('Author').Value)"
        Write-Report "Title: $($Workbook.BuiltinDocumentProperties.Item('Title').Value)"
        Write-Report "Subject: $($Workbook.BuiltinDocumentProperties.Item('Subject').Value)"
        Write-Report "Keywords: $($Workbook.BuiltinDocumentProperties.Item('Keywords').Value)"
        Write-Report "Comments: $($Workbook.BuiltinDocumentProperties.Item('Comments').Value)"
        Write-Report "Created: $($Workbook.BuiltinDocumentProperties.Item('Creation Date').Value)"
    }
    catch {}

    Write-Report ""
    Write-Report "============================================================"
    Write-Report "EXTRACTION COMPLETED"
    Write-Report "============================================================"

    Write-Host ""
    Write-Host "Extraction completed successfully." -ForegroundColor Green
    Write-Host "Output file: $OutputPath" -ForegroundColor Cyan

}
catch {

    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red

    if ($null -ne $script:Writer) {
        Write-Report ""
        Write-Report "EXTRACTION ERROR: $($_.Exception.Message)"
    }
}
finally {

    if ($null -ne $Workbook) {
        $Workbook.Close($false)
    }

    if ($null -ne $Excel) {
        $Excel.Quit()
    }

    if ($null -ne $script:Writer) {
        $script:Writer.Flush()
        $script:Writer.Close()
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