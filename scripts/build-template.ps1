# Build FBM Excel formatting standard template
$ErrorActionPreference = 'Stop'

$outPath = 'C:\ak\fbm-standards\FBM-Excel-Template.xlsx'
$logoPath = 'C:\ak\fbm-standards\fbm-logo.jpeg'
if (Test-Path $outPath) { Remove-Item $outPath -Force }

# Brand colors (BGR for Excel.Interior.Color: 0xBBGGRR)
function RGB($r, $g, $b) { [int]($b * 65536 + $g * 256 + $r) }
$navy      = RGB 9   50  84    # 093254
$burgundy  = RGB 106 24  49    # 6A1831
$slateblue = RGB 78  112 135   # 4E7087
$charcoal  = RGB 99  100 102   # 636466
$forest    = RGB 0   94  52    # 005E34
$sage      = RGB 112 133 115   # 708573
$lightgrey = RGB 231 230 230   # E7E6E6
$white     = RGB 255 255 255
$blueInput = RGB 0   0   255
$black     = RGB 0   0   0
$greenLink = RGB 0   128 0
$redExt    = RGB 255 0   0
$yellow    = RGB 255 255 0
$tabBlue   = RGB 31  78  121   # input tab
$tabGreen  = RGB 0   94  52    # output tab
$tabGray   = RGB 165 165 165   # reference tab

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
    $wb = $excel.Workbooks.Add()
    # Drop default sheets except one we'll rename
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

    # --- Number format strings ---
    $fmtCurrency = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'
    $fmtCurrency2 = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    $fmtNumber = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
    $fmtPercent = '0.0%;(0.0%);"-"'
    $fmtMultiple = '0.0"x"'
    $fmtDate = 'mm/dd/yyyy'

    # --- Helper: create or update a named style ---
    function Set-Style {
        param($name, [scriptblock]$config)
        try { $wb.Styles.Item($name).Delete() } catch {}
        $s = $wb.Styles.Add($name)
        & $config $s
        return $s
    }

    # Inputs (blue text)
    Set-Style 'FBM Input' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$blueInput
        $s.NumberFormat=$fmtNumber
    } | Out-Null

    # Currency input
    Set-Style 'FBM Input $' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$blueInput
        $s.NumberFormat=$fmtCurrency
    } | Out-Null

    # Percent input
    Set-Style 'FBM Input %' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$blueInput
        $s.NumberFormat=$fmtPercent
    } | Out-Null

    # Formula (black text)
    Set-Style 'FBM Formula' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtNumber
    } | Out-Null

    Set-Style 'FBM Formula $' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtCurrency
    } | Out-Null

    Set-Style 'FBM Formula %' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtPercent
    } | Out-Null

    # Cross-sheet link (green text)
    Set-Style 'FBM Link' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$greenLink
        $s.NumberFormat=$fmtNumber
    } | Out-Null

    # External link (red text)
    Set-Style 'FBM External' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$redExt
        $s.NumberFormat=$fmtNumber
    } | Out-Null

    # Key assumption (blue text on yellow fill)
    Set-Style 'FBM Assumption' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$blueInput; $s.Font.Bold=$true
        $s.Interior.Color=$yellow
        $s.NumberFormat='General'
    } | Out-Null

    # Header (navy fill, white bold)
    Set-Style 'FBM Header' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=11; $s.Font.Bold=$true; $s.Font.Color=$white
        $s.Interior.Color=$navy
        $s.HorizontalAlignment = -4108  # xlCenter
        $s.VerticalAlignment = -4108
        $s.Borders.LineStyle = 1
    } | Out-Null

    # Subheader (light grey fill, navy bold)
    Set-Style 'FBM Subheader' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=10; $s.Font.Bold=$true; $s.Font.Color=$navy
        $s.Interior.Color=$lightgrey
        $s.HorizontalAlignment = -4131  # xlLeft
        $s.Borders.LineStyle = 1
    } | Out-Null

    # Title (Rockwell 14 bold, navy)
    Set-Style 'FBM Title' { param($s)
        $s.Font.Name='Rockwell'; $s.Font.Size=14; $s.Font.Bold=$true; $s.Font.Color=$navy
    } | Out-Null

    # Total row (top border, bold)
    Set-Style 'FBM Total' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Bold=$true; $s.Font.Color=$black
        $s.Borders.Item(8).LineStyle = 1   # xlEdgeTop
        $s.Borders.Item(9).LineStyle = 9   # xlEdgeBottom = double
        $s.Borders.Item(9).Weight = 4      # xlThick
        $s.NumberFormat=$fmtNumber
    } | Out-Null

    # Date
    Set-Style 'FBM Date' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtDate
    } | Out-Null

    # Year (text)
    Set-Style 'FBM Year' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat='@'
        $s.HorizontalAlignment = -4108
    } | Out-Null

    # ============================================
    # SHEET 1: COVER
    # ============================================
    $cover = $wb.Sheets.Item(1)
    $cover.Name = 'Cover'
    $cover.Tab.Color = $navy

    # Column widths multiples of 5
    1..10 | ForEach-Object { $cover.Columns.Item($_).ColumnWidth = 15 }
    $cover.Columns.Item(1).ColumnWidth = 5

    # Insert logo
    if (Test-Path $logoPath) {
        $shape = $cover.Shapes.AddPicture($logoPath, $false, $true, 30, 20, 200, 90)
    }

    $cover.Cells.Item(8,2) = 'FBM Excel Formatting Standard'
    $cover.Cells.Item(8,2).Style = 'FBM Title'
    $cover.Cells.Item(8,2).Font.Size = 20

    $cover.Cells.Item(10,2) = '[Workbook Title]'
    $cover.Cells.Item(10,2).Font.Name = 'Calibri'
    $cover.Cells.Item(10,2).Font.Size = 14
    $cover.Cells.Item(10,2).Font.Color = $charcoal

    # Metadata block
    $meta = @(
        @(12,'Prepared by:'),
        @(13,'Department:'),
        @(14,'Date:'),
        @(15,'Version:'),
        @(16,'File location:'),
        @(17,'Purpose:')
    )
    foreach ($row in $meta) {
        $cover.Cells.Item($row[0],2) = $row[1]
        $cover.Cells.Item($row[0],2).Font.Bold = $true
        $cover.Cells.Item($row[0],2).Font.Name = 'Calibri'
        $cover.Cells.Item($row[0],2).Font.Size = 10
        $cover.Cells.Item($row[0],3).Style = 'FBM Input'
    }
    $cover.Cells.Item(12,3) = 'Andrew Kim'
    $cover.Cells.Item(13,3) = 'Operations FP&A'
    $cover.Cells.Item(14,3) = (Get-Date).ToString('MM/dd/yyyy')
    $cover.Cells.Item(14,3).Style = 'FBM Date'; $cover.Cells.Item(14,3).Font.Color = $blueInput
    $cover.Cells.Item(15,3) = '1.0'
    $cover.Cells.Item(17,3) = 'Standard FBM workbook template with cell styles, named ranges, and example sheets.'

    # Legend block
    $cover.Cells.Item(20,2) = 'Legend'
    $cover.Cells.Item(20,2).Style = 'FBM Subheader'
    $cover.Range('B20:D20').Merge()

    $legend = @(
        @(21, 'Cell Style', 'Example', 'Use For'),
        @(22, 'FBM Input', 1000, 'Hardcoded inputs (blue text)'),
        @(23, 'FBM Formula', 2000, 'Calculated formulas (black text)'),
        @(24, 'FBM Link', 3000, 'Links from other sheets (green text)'),
        @(25, 'FBM External', 4000, 'External file links (red text)'),
        @(26, 'FBM Assumption', 0.10, 'Key assumptions (yellow fill, blue bold)')
    )
    # header row
    $cover.Cells.Item(21,2)='Cell Style'; $cover.Cells.Item(21,3)='Example'; $cover.Cells.Item(21,4)='Use For'
    $cover.Range('B21:D21').Style = 'FBM Header'

    $styleNames = @('FBM Input','FBM Formula','FBM Link','FBM External','FBM Assumption')
    $examples   = @(1000, 2000, 3000, 4000, 0.10)
    $descs      = @('Hardcoded inputs (blue text)','Calculated formulas (black text)','Links from other sheets (green text)','External file links (red text)','Key assumption (yellow fill, blue bold)')
    for ($i=0; $i -lt 5; $i++) {
        $r = 22 + $i
        $cover.Cells.Item($r,2) = $styleNames[$i]
        $cover.Cells.Item($r,3) = $examples[$i]
        $cover.Cells.Item($r,3).Style = $styleNames[$i]
        if ($styleNames[$i] -eq 'FBM Assumption') { $cover.Cells.Item($r,3).NumberFormat = $fmtPercent }
        $cover.Cells.Item($r,4) = $descs[$i]
        $cover.Cells.Item($r,4).Font.Name='Calibri'; $cover.Cells.Item($r,4).Font.Size=9
    }

    # Tab color legend
    $cover.Cells.Item(28,2) = 'Tab Color Coding'
    $cover.Cells.Item(28,2).Style = 'FBM Subheader'
    $cover.Range('B28:D28').Merge()
    $cover.Cells.Item(29,2)='Color'; $cover.Cells.Item(29,3)='Purpose'; $cover.Cells.Item(29,4)='Example Sheet'
    $cover.Range('B29:D29').Style = 'FBM Header'
    $cover.Cells.Item(30,2)='Blue';  $cover.Cells.Item(30,2).Interior.Color=$tabBlue;  $cover.Cells.Item(30,2).Font.Color=$white; $cover.Cells.Item(30,2).Font.Bold=$true
    $cover.Cells.Item(30,3)='Inputs / data entry'; $cover.Cells.Item(30,4)='Inputs'
    $cover.Cells.Item(31,2)='Green'; $cover.Cells.Item(31,2).Interior.Color=$tabGreen; $cover.Cells.Item(31,2).Font.Color=$white; $cover.Cells.Item(31,2).Font.Bold=$true
    $cover.Cells.Item(31,3)='Outputs / results / summaries'; $cover.Cells.Item(31,4)='Output'
    $cover.Cells.Item(32,2)='Gray';  $cover.Cells.Item(32,2).Interior.Color=$tabGray;  $cover.Cells.Item(32,2).Font.Color=$white; $cover.Cells.Item(32,2).Font.Bold=$true
    $cover.Cells.Item(32,3)='Reference / lookup tables'; $cover.Cells.Item(32,4)='Reference'

    # Page setup for Cover
    $cover.PageSetup.Orientation = 1  # xlPortrait
    $cover.PageSetup.LeftMargin   = $excel.InchesToPoints(0.5)
    $cover.PageSetup.RightMargin  = $excel.InchesToPoints(0.5)
    $cover.PageSetup.TopMargin    = $excel.InchesToPoints(0.5)
    $cover.PageSetup.BottomMargin = $excel.InchesToPoints(0.5)
    $cover.PageSetup.CenterHeader = '&"Calibri,Bold"&14FBM Excel Formatting Standard'
    $cover.PageSetup.RightFooter = '&P of &N'
    $cover.PageSetup.LeftFooter = '&F'

    # ============================================
    # SHEET 2: INPUTS (Blue tab)
    # ============================================
    $inputs = $wb.Sheets.Add([System.Reflection.Missing]::Value, $cover)
    $inputs.Name = 'Inputs'
    $inputs.Tab.Color = $tabBlue

    $inputs.Columns.Item(1).ColumnWidth = 5
    2..10 | ForEach-Object { $inputs.Columns.Item($_).ColumnWidth = 15 }
    $inputs.Columns.Item(2).ColumnWidth = 30

    $inputs.Cells.Item(2,2) = 'Inputs'
    $inputs.Cells.Item(2,2).Style = 'FBM Title'
    $inputs.Cells.Item(3,2) = 'Blue text cells are user-editable. Do not overwrite black formula cells.'
    $inputs.Cells.Item(3,2).Font.Italic = $true
    $inputs.Cells.Item(3,2).Font.Color = $charcoal
    $inputs.Cells.Item(3,2).Font.Size = 9
    # Subtitle: wrap OFF so long sentences overflow right instead of inflating row height
    $inputs.Cells.Item(3,2).WrapText = $false
    $inputs.Cells.Item(3,2).HorizontalAlignment = -4131  # xlLeft

    # Header row
    $inputs.Range('B5:F5').Style = 'FBM Header'
    $inputs.Cells.Item(5,2)='Item'
    $inputs.Cells.Item(5,3)='FY2026'
    $inputs.Cells.Item(5,4)='FY2027'
    $inputs.Cells.Item(5,5)='FY2028'
    $inputs.Cells.Item(5,6)='FY2029'
    # Year cells use text format
    $inputs.Range('C5:F5').HorizontalAlignment = -4108

    # Sample input rows
    $items = @('Revenue ($)', 'Growth rate (%)', 'COGS ($)', 'Headcount (#)')
    $styles = @('FBM Input $', 'FBM Input %', 'FBM Input $', 'FBM Input')
    $vals = @(
        @(10000000, 10500000, 11025000, 11576250),
        @(0.05, 0.05, 0.05, 0.05),
        @(6000000, 6300000, 6615000, 6945750),
        @(50, 52, 55, 58)
    )
    for ($i=0; $i -lt 4; $i++) {
        $r = 6 + $i
        $inputs.Cells.Item($r,2) = $items[$i]
        $inputs.Cells.Item($r,2).Font.Name='Calibri'; $inputs.Cells.Item($r,2).Font.Size=9
        for ($c=0; $c -lt 4; $c++) {
            $inputs.Cells.Item($r, 3+$c) = $vals[$i][$c]
            $inputs.Cells.Item($r, 3+$c).Style = $styles[$i]
        }
    }

    # Named range example - Inputs!Revenue refers to row 6
    $wb.Names.Add('Revenue_Input', "=Inputs!`$C`$6:`$F`$6") | Out-Null
    $wb.Names.Add('Growth_Rate', "=Inputs!`$C`$7:`$F`$7") | Out-Null

    # Freeze panes: one row below header_row (5) and one col right of label_col (B) -> C6.
    # When the header sits elsewhere, change the Select() target accordingly.
    $inputs.Activate()
    $inputs.Range('C6').Select()
    $excel.ActiveWindow.FreezePanes = $true

    # Auto filter
    $inputs.Range('B5:F5').AutoFilter() | Out-Null

    # Page setup
    $inputs.PageSetup.Orientation = 2  # landscape
    $inputs.PageSetup.Zoom = $false
    $inputs.PageSetup.FitToPagesWide = 1
    $inputs.PageSetup.FitToPagesTall = $false
    $inputs.PageSetup.LeftMargin   = $excel.InchesToPoints(0.5)
    $inputs.PageSetup.RightMargin  = $excel.InchesToPoints(0.5)
    $inputs.PageSetup.TopMargin    = $excel.InchesToPoints(0.75)
    $inputs.PageSetup.BottomMargin = $excel.InchesToPoints(0.5)
    $inputs.PageSetup.PrintTitleRows = '$5:$5'
    $inputs.PageSetup.CenterHeader = '&"Calibri,Bold"&12Inputs'
    $inputs.PageSetup.LeftHeader = '&G'  # placeholder - we'll add picture below
    $inputs.PageSetup.RightFooter = '&P of &N'
    $inputs.PageSetup.LeftFooter = '&F'
    $inputs.PageSetup.CenterFooter = '&D'

    # ============================================
    # SHEET 3: CALC (no special tab color)
    # ============================================
    $calc = $wb.Sheets.Add([System.Reflection.Missing]::Value, $inputs)
    $calc.Name = 'Calc'

    $calc.Columns.Item(1).ColumnWidth = 5
    2..10 | ForEach-Object { $calc.Columns.Item($_).ColumnWidth = 15 }
    $calc.Columns.Item(2).ColumnWidth = 30

    $calc.Cells.Item(2,2) = 'Calculations'
    $calc.Cells.Item(2,2).Style = 'FBM Title'

    $calc.Range('B5:F5').Style = 'FBM Header'
    $calc.Cells.Item(5,2)='Metric'
    $calc.Cells.Item(5,3)='FY2026'
    $calc.Cells.Item(5,4)='FY2027'
    $calc.Cells.Item(5,5)='FY2028'
    $calc.Cells.Item(5,6)='FY2029'
    $calc.Range('C5:F5').HorizontalAlignment = -4108

    # Link rows from Inputs (green)
    $calc.Cells.Item(6,2)='Revenue'
    $calc.Cells.Item(6,2).Font.Name='Calibri'; $calc.Cells.Item(6,2).Font.Size=9
    $calc.Range('C6:F6').Style = 'FBM Link'
    $calc.Range('C6:F6').NumberFormat = $fmtCurrency
    $calc.Cells.Item(6,3).Formula = '=Inputs!C6'
    $calc.Cells.Item(6,4).Formula = '=Inputs!D6'
    $calc.Cells.Item(6,5).Formula = '=Inputs!E6'
    $calc.Cells.Item(6,6).Formula = '=Inputs!F6'

    $calc.Cells.Item(7,2)='COGS'
    $calc.Cells.Item(7,2).Font.Name='Calibri'; $calc.Cells.Item(7,2).Font.Size=9
    $calc.Range('C7:F7').Style = 'FBM Link'
    $calc.Range('C7:F7').NumberFormat = $fmtCurrency
    $calc.Cells.Item(7,3).Formula = '=Inputs!C8'
    $calc.Cells.Item(7,4).Formula = '=Inputs!D8'
    $calc.Cells.Item(7,5).Formula = '=Inputs!E8'
    $calc.Cells.Item(7,6).Formula = '=Inputs!F8'

    # Calculation row (black formula)
    $calc.Cells.Item(8,2)='Gross Profit'
    $calc.Cells.Item(8,2).Font.Name='Calibri'; $calc.Cells.Item(8,2).Font.Size=9
    $calc.Range('C8:F8').Style = 'FBM Formula $'
    $calc.Cells.Item(8,3).Formula = '=C6-C7'
    $calc.Cells.Item(8,4).Formula = '=D6-D7'
    $calc.Cells.Item(8,5).Formula = '=E6-E7'
    $calc.Cells.Item(8,6).Formula = '=F6-F7'

    # Gross margin %
    $calc.Cells.Item(9,2)='Gross Margin %'
    $calc.Cells.Item(9,2).Font.Name='Calibri'; $calc.Cells.Item(9,2).Font.Size=9
    $calc.Range('C9:F9').Style = 'FBM Formula %'
    $calc.Cells.Item(9,3).Formula = '=IFERROR(C8/C6,0)'
    $calc.Cells.Item(9,4).Formula = '=IFERROR(D8/D6,0)'
    $calc.Cells.Item(9,5).Formula = '=IFERROR(E8/E6,0)'
    $calc.Cells.Item(9,6).Formula = '=IFERROR(F8/F6,0)'

    # Conditional formatting on gross margin: good/bad/neutral
    $cfRange = $calc.Range('C9:F9')
    $cf = $cfRange.FormatConditions.AddIconSetCondition()
    $cf.IconSet = $wb.IconSets.Item(1) # xl3TrafficLights1 = 1
    $cf.IconCriteria.Item(2).Type = 0  # xlConditionValueNumber
    $cf.IconCriteria.Item(2).Value = 0.3
    $cf.IconCriteria.Item(3).Type = 0
    $cf.IconCriteria.Item(3).Value = 0.4

    $calc.Activate()
    $calc.Range('C6').Select()
    $excel.ActiveWindow.FreezePanes = $true

    $calc.PageSetup.Orientation = 2
    $calc.PageSetup.Zoom = $false
    $calc.PageSetup.FitToPagesWide = 1
    $calc.PageSetup.FitToPagesTall = $false
    $calc.PageSetup.PrintTitleRows = '$5:$5'
    $calc.PageSetup.CenterHeader = '&"Calibri,Bold"&12Calculations'
    $calc.PageSetup.RightFooter = '&P of &N'
    $calc.PageSetup.LeftFooter = '&F'

    # ============================================
    # SHEET 4: OUTPUT (Green tab)
    # ============================================
    $output = $wb.Sheets.Add([System.Reflection.Missing]::Value, $calc)
    $output.Name = 'Output'
    $output.Tab.Color = $tabGreen

    $output.Columns.Item(1).ColumnWidth = 5
    2..10 | ForEach-Object { $output.Columns.Item($_).ColumnWidth = 15 }
    $output.Columns.Item(2).ColumnWidth = 30

    $output.Cells.Item(2,2) = 'Output Summary'
    $output.Cells.Item(2,2).Style = 'FBM Title'

    $output.Range('B5:F5').Style = 'FBM Header'
    $output.Cells.Item(5,2)='KPI'
    $output.Cells.Item(5,3)='FY2026'
    $output.Cells.Item(5,4)='FY2027'
    $output.Cells.Item(5,5)='FY2028'
    $output.Cells.Item(5,6)='FY2029'
    $output.Range('C5:F5').HorizontalAlignment = -4108

    $output.Cells.Item(6,2)='Revenue'
    $output.Cells.Item(6,2).Font.Name='Calibri'; $output.Cells.Item(6,2).Font.Size=9
    $output.Range('C6:F6').Style = 'FBM Link'
    $output.Range('C6:F6').NumberFormat = $fmtCurrency
    $output.Cells.Item(6,3).Formula = '=Calc!C6'
    $output.Cells.Item(6,4).Formula = '=Calc!D6'
    $output.Cells.Item(6,5).Formula = '=Calc!E6'
    $output.Cells.Item(6,6).Formula = '=Calc!F6'

    $output.Cells.Item(7,2)='Gross Profit'
    $output.Cells.Item(7,2).Font.Name='Calibri'; $output.Cells.Item(7,2).Font.Size=9
    $output.Range('C7:F7').Style = 'FBM Link'
    $output.Range('C7:F7').NumberFormat = $fmtCurrency
    $output.Cells.Item(7,3).Formula = '=Calc!C8'
    $output.Cells.Item(7,4).Formula = '=Calc!D8'
    $output.Cells.Item(7,5).Formula = '=Calc!E8'
    $output.Cells.Item(7,6).Formula = '=Calc!F8'

    $output.Cells.Item(8,2)='Gross Margin %'
    $output.Cells.Item(8,2).Font.Name='Calibri'; $output.Cells.Item(8,2).Font.Size=9
    $output.Range('C8:F8').Style = 'FBM Link'
    $output.Range('C8:F8').NumberFormat = $fmtPercent
    $output.Cells.Item(8,3).Formula = '=Calc!C9'
    $output.Cells.Item(8,4).Formula = '=Calc!D9'
    $output.Cells.Item(8,5).Formula = '=Calc!E9'
    $output.Cells.Item(8,6).Formula = '=Calc!F9'

    $output.Cells.Item(10,2) = 'Total Revenue (4Y)'
    $output.Cells.Item(10,2).Font.Bold = $true
    $output.Cells.Item(10,2).Font.Name='Calibri'; $output.Cells.Item(10,2).Font.Size=9
    $output.Cells.Item(10,3).Formula = '=SUM(C6:F6)'
    $output.Cells.Item(10,3).Style = 'FBM Total'
    $output.Cells.Item(10,3).NumberFormat = $fmtCurrency

    $output.Activate()
    $output.Range('C6').Select()
    $excel.ActiveWindow.FreezePanes = $true

    $output.PageSetup.Orientation = 2
    $output.PageSetup.Zoom = $false
    $output.PageSetup.FitToPagesWide = 1
    $output.PageSetup.FitToPagesTall = $false
    $output.PageSetup.PrintTitleRows = '$5:$5'
    $output.PageSetup.CenterHeader = '&"Calibri,Bold"&12Output Summary'
    $output.PageSetup.RightFooter = '&P of &N'
    $output.PageSetup.LeftFooter = '&F'

    # ============================================
    # SHEET 5: REFERENCE (Gray tab)
    # ============================================
    $ref = $wb.Sheets.Add([System.Reflection.Missing]::Value, $output)
    $ref.Name = 'Reference'
    $ref.Tab.Color = $tabGray

    $ref.Columns.Item(1).ColumnWidth = 5
    2..10 | ForEach-Object { $ref.Columns.Item($_).ColumnWidth = 15 }
    $ref.Columns.Item(2).ColumnWidth = 30
    $ref.Columns.Item(3).ColumnWidth = 25

    $ref.Cells.Item(2,2) = 'Reference / Lookup Tables'
    $ref.Cells.Item(2,2).Style = 'FBM Title'

    $ref.Range('B5:C5').Style = 'FBM Header'
    $ref.Cells.Item(5,2)='Key'
    $ref.Cells.Item(5,3)='Value'

    # Example reference data
    $refData = @(
        @('Discount rate', 0.10),
        @('Tax rate', 0.21),
        @('FY start month', 1),
        @('Working days/yr', 252)
    )
    for ($i=0; $i -lt $refData.Count; $i++) {
        $r = 6 + $i
        $ref.Cells.Item($r,2) = $refData[$i][0]
        $ref.Cells.Item($r,2).Font.Name='Calibri'; $ref.Cells.Item($r,2).Font.Size=9
        $ref.Cells.Item($r,3) = $refData[$i][1]
        $ref.Cells.Item($r,3).Style = 'FBM Input'
        if ($refData[$i][0] -like '*rate*') { $ref.Cells.Item($r,3).NumberFormat = $fmtPercent }
    }

    # Add named ranges for lookups
    $wb.Names.Add('Discount_Rate', "=Reference!`$C`$6") | Out-Null
    $wb.Names.Add('Tax_Rate',      "=Reference!`$C`$7") | Out-Null

    $ref.Activate()
    $ref.Range('C6').Select()
    $excel.ActiveWindow.FreezePanes = $true

    $ref.PageSetup.Orientation = 1
    $ref.PageSetup.CenterHeader = '&"Calibri,Bold"&12Reference'
    $ref.PageSetup.RightFooter = '&P of &N'
    $ref.PageSetup.LeftFooter = '&F'

    # Make Cover the active sheet on open
    $cover.Activate()

    # Save
    # xlOpenXMLWorkbook = 51
    $wb.SaveAs($outPath, 51)
    $wb.Close($false)
    Write-Output "SAVED: $outPath"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
