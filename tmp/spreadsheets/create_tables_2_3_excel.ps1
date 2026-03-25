$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$outputDir = Join-Path $root 'output\spreadsheet'
$tmpDir = Join-Path $root 'tmp\spreadsheets'
$workbookPath = Join-Path $outputDir 'tables_2_3_editable.xlsx'
$previewPdfPath = Join-Path $tmpDir 'tables_2_3_editable_preview.pdf'

function Set-CellValue {
    param(
        [Parameter(Mandatory = $true)] $Sheet,
        [Parameter(Mandatory = $true)] [string] $Address,
        [Parameter(Mandatory = $true)] $Value
    )
    $Sheet.Range($Address).Value2 = $Value
}

function Apply-TitleStyle {
    param(
        [Parameter(Mandatory = $true)] $Range
    )
    $Range.Font.Bold = $true
    $Range.Font.Size = 14
    $Range.Interior.Color = 15773696
}

function Apply-HeaderStyle {
    param(
        [Parameter(Mandatory = $true)] $Range
    )
    $Range.Font.Bold = $true
    $Range.WrapText = $true
    $Range.HorizontalAlignment = -4108
    $Range.VerticalAlignment = -4108
    $Range.Interior.Color = 14737632
}

function Apply-InputStyle {
    param(
        [Parameter(Mandatory = $true)] $Range
    )
    $Range.Interior.Color = 15773696
}

function Apply-BorderStyle {
    param(
        [Parameter(Mandatory = $true)] $Range
    )
    $Range.Borders.LineStyle = 1
    $Range.Borders.Weight = 2
}

function Configure-PageLayout {
    param(
        [Parameter(Mandatory = $true)] $Sheet,
        [Parameter(Mandatory = $true)] [string] $PrintArea,
        [Parameter(Mandatory = $false)] [int] $Orientation = 2
    )
    $Sheet.PageSetup.PrintArea = $PrintArea
    $Sheet.PageSetup.Zoom = $false
    $Sheet.PageSetup.FitToPagesWide = 1
    $Sheet.PageSetup.FitToPagesTall = 1
    $Sheet.PageSetup.Orientation = $Orientation
    $Sheet.PageSetup.CenterHorizontally = $true
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    while ($workbook.Worksheets.Count -lt 3) {
        [void]$workbook.Worksheets.Add()
    }
    while ($workbook.Worksheets.Count -gt 3) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $instructions = $workbook.Worksheets.Item(1)
    $table2 = $workbook.Worksheets.Item(2)
    $table3 = $workbook.Worksheets.Item(3)

    $instructions.Name = 'Instructions'
    $table2.Name = 'Table 2'
    $table3.Name = 'Table 3'

    # Instructions sheet
    Set-CellValue $instructions 'A1' 'Editable workbook for report Table 2 and Table 3'
    Set-CellValue $instructions 'A3' 'Blue cells are intended for editable numeric inputs.'
    Set-CellValue $instructions 'A4' 'Percentage cells update automatically from formulas.'
    Set-CellValue $instructions 'A5' 'Table 2 percentages are absolute change relative to Mesh 3.'
    Set-CellValue $instructions 'A6' 'Table 3 relative change is calculated as (Full Frame - Arch Only) / Arch Only.'
    Set-CellValue $instructions 'A7' 'For non-comparable quantities in Table 3, leave Arch Only blank or set it to 0.'
    Apply-TitleStyle $instructions.Range('A1:F1')
    $instructions.Range('A1:F1').Merge()
    $instructions.Range('A3:A7').Font.Size = 11
    $instructions.Columns('A:F').ColumnWidth = 22
    Configure-PageLayout $instructions '$A$1:$F$7' 1

    # Table 2 sheet
    Set-CellValue $table2 'A1' 'Table 2: Influence of Arch Mesh Density on Key Global Responses'
    Set-CellValue $table2 'A2' 'Source values taken from the current report table. Edit the blue cells as needed.'
    Apply-TitleStyle $table2.Range('A1:K1')
    $table2.Range('A1:K1').Merge()
    $table2.Range('A2:K2').Merge()
    $table2.Range('A3').Value2 = 'Mesh'
    $table2.Range('B3').Value2 = 'Arch Elements'
    $table2.Range('C3').Value2 = 'DoF'
    $table2.Range('D3').Value2 = 'Arch Defl. (mm)'
    $table2.Range('E3').Value2 = 'Abs Change vs Mesh 3'
    $table2.Range('F3').Value2 = 'Tie Force (kN)'
    $table2.Range('G3').Value2 = 'Abs Change vs Mesh 3'
    $table2.Range('H3').Value2 = 'Arch Mid BM (kNm)'
    $table2.Range('I3').Value2 = 'Abs Change vs Mesh 3'
    $table2.Range('J3').Value2 = 'Tusk-Arch BM (kNm)'
    $table2.Range('K3').Value2 = 'Abs Change vs Mesh 3'
    Apply-HeaderStyle $table2.Range('A3:K3')

    $table2Data = @(
        @('Mesh 1', 14, 53, 241.2, 9725, 12760, 4721),
        @('Mesh 2', 16, 59, 175.7, 9309, 960.2, 5753),
        @('Mesh 3', 22, 77, 175.3, 9320, 198.2, 5837)
    )

    $row = 4
    foreach ($entry in $table2Data) {
        $table2.Range('A' + $row).Value2 = [string]$entry[0]
        $table2.Range('B' + $row).Value2 = [double]$entry[1]
        $table2.Range('C' + $row).Value2 = [double]$entry[2]
        $table2.Range('D' + $row).Value2 = [double]$entry[3]
        $table2.Range('F' + $row).Value2 = [double]$entry[4]
        $table2.Range('H' + $row).Value2 = [double]$entry[5]
        $table2.Range('J' + $row).Value2 = [double]$entry[6]
        $row++
    }

    $table2.Range('E4').Formula = '=IFERROR(ABS((D4-$D$6)/$D$6),"")'
    $table2.Range('E5').Formula = '=IFERROR(ABS((D5-$D$6)/$D$6),"")'
    $table2.Range('E6').Formula = '=IFERROR(ABS((D6-$D$6)/$D$6),"")'
    $table2.Range('G4').Formula = '=IFERROR(ABS((F4-$F$6)/$F$6),"")'
    $table2.Range('G5').Formula = '=IFERROR(ABS((F5-$F$6)/$F$6),"")'
    $table2.Range('G6').Formula = '=IFERROR(ABS((F6-$F$6)/$F$6),"")'
    $table2.Range('I4').Formula = '=IFERROR(ABS((H4-$H$6)/$H$6),"")'
    $table2.Range('I5').Formula = '=IFERROR(ABS((H5-$H$6)/$H$6),"")'
    $table2.Range('I6').Formula = '=IFERROR(ABS((H6-$H$6)/$H$6),"")'
    $table2.Range('K4').Formula = '=IFERROR(ABS((J4-$J$6)/$J$6),"")'
    $table2.Range('K5').Formula = '=IFERROR(ABS((J5-$J$6)/$J$6),"")'
    $table2.Range('K6').Formula = '=IFERROR(ABS((J6-$J$6)/$J$6),"")'

    Apply-InputStyle $table2.Range('B4:D6')
    Apply-InputStyle $table2.Range('F4:F6')
    Apply-InputStyle $table2.Range('H4:H6')
    Apply-InputStyle $table2.Range('J4:J6')
    Apply-BorderStyle $table2.Range('A3:K6')

    $table2.Range('B4:C6').NumberFormat = '0'
    $table2.Range('D4:D6').NumberFormat = '0.0'
    $table2.Range('F4:F6').NumberFormat = '0'
    $table2.Range('H4:H6').NumberFormat = '0.0'
    $table2.Range('J4:J6').NumberFormat = '0'
    $table2.Range('E4:E6').NumberFormat = '0.0%'
    $table2.Range('G4:G6').NumberFormat = '0.0%'
    $table2.Range('I4:I6').NumberFormat = '0.0%'
    $table2.Range('K4:K6').NumberFormat = '0.0%'
    $table2.Range('A3:K6').VerticalAlignment = -4108
    $table2.Range('A3:K6').WrapText = $true
    $table2.Columns('A').ColumnWidth = 14
    $table2.Columns('B').ColumnWidth = 13
    $table2.Columns('C').ColumnWidth = 9
    $table2.Columns('D').ColumnWidth = 14
    $table2.Columns('E').ColumnWidth = 17
    $table2.Columns('F').ColumnWidth = 13
    $table2.Columns('G').ColumnWidth = 17
    $table2.Columns('H').ColumnWidth = 15
    $table2.Columns('I').ColumnWidth = 17
    $table2.Columns('J').ColumnWidth = 15
    $table2.Columns('K').ColumnWidth = 17
    $table2.Rows('1:6').RowHeight = 24
    $table2.Application.ActiveWindow.SplitRow = 3
    $table2.Application.ActiveWindow.FreezePanes = $true
    Configure-PageLayout $table2 '$A$1:$K$6' 2

    # Table 3 sheet
    Set-CellValue $table3 'A1' 'Table 3: Comparison Between Prestressing Assumptions'
    Set-CellValue $table3 'A2' 'Relative change updates automatically when both model values are numeric and Arch Only is non-zero.'
    Apply-TitleStyle $table3.Range('A1:E1')
    $table3.Range('A1:E1').Merge()
    $table3.Range('A2:E2').Merge()
    $table3.Range('A3').Value2 = 'Response Quantity'
    $table3.Range('B3').Value2 = 'Arch Only Before Lift'
    $table3.Range('C3').Value2 = 'Full Frame After Erection'
    $table3.Range('D3').Value2 = 'Relative Change'
    $table3.Range('E3').Value2 = 'Notes'
    Apply-HeaderStyle $table3.Range('A3:E3')

    $table3Rows = @(
        @('Maximum vertical displacement (mm)', 213.3, 175.3, ''),
        @('Central arch bending moment (kNm)', 4128, 198.2, ''),
        @('Tusk-arch connection bending moment (kNm)', 0, 5837, 'Currently treated as not comparable'),
        @('High tie force (kN)', 11450, 9320, ''),
        @('Horizontal reaction at supports (kN)', '', 3480, 'Arch-only value not available in current report')
    )

    $row = 4
    foreach ($entry in $table3Rows) {
        $table3.Range('A' + $row).Value2 = [string]$entry[0]
        if ($entry[1] -ne '') {
            $table3.Range('B' + $row).Value2 = [double]$entry[1]
        }
        if ($entry[2] -ne '') {
            $table3.Range('C' + $row).Value2 = [double]$entry[2]
        }
        $table3.Range('E' + $row).Value2 = [string]$entry[3]
        $table3.Cells.Item($row, 4).Formula = '=IF(OR(B' + $row + '="",C' + $row + '="",B' + $row + '=0),"Not comparable",(C' + $row + '-B' + $row + ')/B' + $row + ')'
        $row++
    }

    Apply-InputStyle $table3.Range('B4:C8')
    Apply-BorderStyle $table3.Range('A3:E8')

    $table3.Range('B4:C8').NumberFormat = '0.0'
    $table3.Range('D4:D8').NumberFormat = '0.0%'
    $table3.Range('A3:E8').VerticalAlignment = -4108
    $table3.Range('A3:E8').WrapText = $true
    $table3.Columns('A').ColumnWidth = 34
    $table3.Columns('B').ColumnWidth = 18
    $table3.Columns('C').ColumnWidth = 20
    $table3.Columns('D').ColumnWidth = 16
    $table3.Columns('E').ColumnWidth = 30
    $table3.Rows('1:8').RowHeight = 24
    Configure-PageLayout $table3 '$A$1:$E$8' 2

    $workbook.Worksheets.Item('Instructions').Activate() | Out-Null

    $workbook.SaveAs($workbookPath, 51)
    $workbook.ExportAsFixedFormat(0, $previewPdfPath)
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($true)
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
