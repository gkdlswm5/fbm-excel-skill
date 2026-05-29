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
$bandGrey  = RGB 247 247 247   # F7F7F7 (alternating data-row fill)
$white     = RGB 255 255 255
$blueInput = RGB 0   0   255
$black     = RGB 0   0   0
$greenLink = RGB 0   128 0
$redExt    = RGB 255 0   0
$yellow    = RGB 255 255 0
$tabBlue   = RGB 31  78  121   # input tab
$tabGreen  = RGB 0   94  52    # output tab
$tabGray   = RGB 165 165 165   # reference tab

# Workbook metadata (for properties + Cover sheet)
$workbookTitle    = 'FBM Excel Formatting Standard'
$workbookSubject  = 'Working template with cell styles, named ranges, example sheets'
$workbookAuthor   = 'Andrew Kim'
$workbookCompany  = 'Foundation Building Materials'
$workbookCategory = 'Operations FP&A'
$workbookVersion  = '1.2'
$asOfDate         = (Get-Date).ToString('MM/dd/yyyy')
$reportingScale   = '$ thousands'
$reportingCcy     = 'USD'
$contactEmail     = 'andrew.kim@fbm.com'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
    $wb = $excel.Workbooks.Add()
    # Drop default sheets except one we'll rename
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

    # --- Number format strings ---
    $fmtCurrency  = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'
    $fmtCurrency2 = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    $fmtNumber    = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
    $fmtPercent   = '0.0%;(0.0%);"-"'
    $fmtMultiple  = '0.0"x"'
    $fmtDate      = 'mm/dd/yyyy'
    $fmtRatio     = '0.00'
    $fmtBps       = '0" bps"'
    $fmtShareCt   = '#,##0'
    $fmtSharePx   = '$#,##0.00'
    # Variance formats - "positive good, negative bad"
    $fmtVar       = '+#,##0;(#,##0);"-"'
    $fmtVarDol    = '+$#,##0;($#,##0);"-"'
    $fmtVarPct    = '+0.0%;(0.0%);"-"'
    $fmtVarBps    = '+0" bps";(0)" bps";"-"'

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

    # Subtitle (italic charcoal, wrap OFF, left aligned)
    Set-Style 'FBM Subtitle' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Italic=$true; $s.Font.Color=$charcoal
        $s.HorizontalAlignment = -4131  # xlLeft
        $s.VerticalAlignment   = -4108  # xlCenter
        $s.WrapText = $false
    } | Out-Null

    # Units (italic charcoal, centered, wrap OFF)
    Set-Style 'FBM Units' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Italic=$true; $s.Font.Color=$charcoal
        $s.HorizontalAlignment = -4108  # xlCenter
        $s.VerticalAlignment   = -4108
        $s.WrapText = $false
    } | Out-Null

    # Variance styles - positive good, negative bad
    Set-Style 'FBM Variance' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtVar
    } | Out-Null

    Set-Style 'FBM Variance $' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtVarDol
    } | Out-Null

    Set-Style 'FBM Variance %' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtVarPct
    } | Out-Null

    Set-Style 'FBM Variance bps' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.NumberFormat=$fmtVarBps
    } | Out-Null

    # Banding (alternating data-row fill)
    Set-Style 'FBM Band' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Color=$black
        $s.Interior.Color=$bandGrey
    } | Out-Null

    # KPI Big (24pt Rockwell bold navy, centered)
    Set-Style 'FBM KPI Big' { param($s)
        $s.Font.Name='Rockwell'; $s.Font.Size=24; $s.Font.Bold=$true; $s.Font.Color=$navy
        $s.HorizontalAlignment = -4108
        $s.VerticalAlignment   = -4108
    } | Out-Null

    # KPI Label (9pt italic charcoal, centered)
    Set-Style 'FBM KPI Label' { param($s)
        $s.Font.Name='Calibri'; $s.Font.Size=9; $s.Font.Italic=$true; $s.Font.Color=$charcoal
        $s.HorizontalAlignment = -4108
        $s.VerticalAlignment   = -4108
    } | Out-Null

    # --- Helper: apply banding via conditional format to a range ---
    function Set-Banding {
        param($range)
        $cf = $range.FormatConditions.Add(2, 0, '=MOD(ROW(),2)=0')  # xlExpression
        $cf.Interior.Color = $bandGrey
    }

    # --- Helper: apply standard row heights ---
    function Set-RowHeights {
        param($ws, [int]$headerRow = 5, [bool]$hasSubtitle = $true, [bool]$hasUnits = $false)
        $ws.Rows.Item(2).RowHeight = 24
        if ($hasSubtitle) { $ws.Rows.Item(3).RowHeight = 16 }
        if ($hasUnits)    { $ws.Rows.Item(4).RowHeight = 16 }
        $ws.Rows.Item($headerRow).RowHeight = 30
        for ($r = $headerRow + 1; $r -le $headerRow + 20; $r++) {
            $ws.Rows.Item($r).RowHeight = 16
        }
    }

    # --- Helper: standard print footer (file path / date+version+as-of / page) ---
    function Set-FbmFooter {
        param($ws)
        $ws.PageSetup.LeftFooter   = '&F'
        $ws.PageSetup.CenterFooter = "&D  $([char]0x00B7)  v$workbookVersion  $([char]0x00B7)  As of $asOfDate"
        $ws.PageSetup.RightFooter  = '&P of &N'
    }

    # ============================================
    # SHEET 1: COVER
    # ============================================
    $cover = $wb.Sheets.Item(1)
    $cover.Name = 'Cover'
    $cover.Tab.Color = $navy

    # Column widths multiples of 5
    1..10 | ForEach-Object { $cover.Columns.Item($_).ColumnWidth = 15 }
    $cover.Columns.Item(1).ColumnWidth = 5
    $cover.Columns.Item(2).ColumnWidth = 30

    # Insert logo
    if (Test-Path $logoPath) {
        $shape = $cover.Shapes.AddPicture($logoPath, $false, $true, 30, 20, 200, 90)
    }

    # Title
    $cover.Cells.Item(8,2) = $workbookTitle
    $cover.Cells.Item(8,2).Style = 'FBM Title'
    $cover.Cells.Item(8,2).Font.Size = 20
    $cover.Rows.Item(8).RowHeight = 30

    # Subtitle
    $cover.Cells.Item(10,2) = '[Workbook Title]'
    $cover.Cells.Item(10,2).Font.Name = 'Calibri'
    $cover.Cells.Item(10,2).Font.Size = 14
    $cover.Cells.Item(10,2).Font.Color = $charcoal

    # Metadata block (rows 12-20)
    $metaRows = @(
        @(12, 'Prepared by:',         $workbookAuthor),
        @(13, 'Department:',          'Operations FP&A'),
        @(14, 'As-of date:',          $asOfDate),
        @(15, 'Reporting scale:',     $reportingScale),
        @(16, 'Reporting currency:',  $reportingCcy),
        @(17, 'Contact:',             $contactEmail),
        @(18, 'Version:',             $workbookVersion),
        @(19, 'File location:',       ''),
        @(20, 'Purpose:',             $workbookSubject)
    )
    foreach ($row in $metaRows) {
        $r = $row[0]
        $cover.Cells.Item($r, 2) = $row[1]
        $cover.Cells.Item($r, 2).Font.Bold = $true
        $cover.Cells.Item($r, 2).Font.Name = 'Calibri'
        $cover.Cells.Item($r, 2).Font.Size = 10
        $cover.Cells.Item($r, 3) = $row[2]
        $cover.Cells.Item($r, 3).Style = 'FBM Input'
    }
    # Mark As-of date as a date cell
    $cover.Cells.Item(14, 3).Style = 'FBM Date'
    $cover.Cells.Item(14, 3).Font.Color = $blueInput

    # Named ranges anchored to the metadata block
    $wb.Names.Add('AsOfDate',          "=Cover!`$C`$14") | Out-Null
    $wb.Names.Add('Scale',             "=Cover!`$C`$15") | Out-Null
    $wb.Names.Add('ReportingCurrency', "=Cover!`$C`$16") | Out-Null

    # Table of Contents (rows 22-27)
    $cover.Cells.Item(22, 2) = 'Table of Contents'
    $cover.Cells.Item(22, 2).Style = 'FBM Subheader'
    $cover.Range('B22:D22').Merge()
    $tocSheets = @('Inputs', 'Calc', 'Output', 'Reference')
    for ($i = 0; $i -lt $tocSheets.Count; $i++) {
        $r = 23 + $i
        $name = $tocSheets[$i]
        $cover.Cells.Item($r, 2).Formula = "=HYPERLINK(""#'$name'!A1"",""$name"")"
        $cover.Cells.Item($r, 2).Font.Name  = 'Calibri'
        $cover.Cells.Item($r, 2).Font.Size  = 9
        $cover.Cells.Item($r, 2).Font.Color = $greenLink
    }

    # Data Sources block (rows 29-33)
    $cover.Cells.Item(29, 2) = 'Data Sources'
    $cover.Cells.Item(29, 2).Style = 'FBM Subheader'
    $cover.Range('B29:E29').Merge()
    $cover.Cells.Item(30, 2) = 'Source system';   $cover.Cells.Item(30, 3) = 'Last refreshed'; $cover.Cells.Item(30, 4) = 'Owner';  $cover.Cells.Item(30, 5) = 'Notes'
    $cover.Range('B30:E30').Style = 'FBM Header'
    $cover.Cells.Item(31, 2) = '[e.g. SAP S/4HANA]'; $cover.Cells.Item(31, 3) = '[YYYY-MM-DD]'; $cover.Cells.Item(31, 4) = '[name]'; $cover.Cells.Item(31, 5) = '[notes]'
    $cover.Range('B31:E31').Font.Name = 'Calibri'; $cover.Range('B31:E31').Font.Size = 9

    # Version History block (rows 35-40)
    $cover.Cells.Item(35, 2) = 'Version History'
    $cover.Cells.Item(35, 2).Style = 'FBM Subheader'
    $cover.Range('B35:E35').Merge()
    $cover.Cells.Item(36, 2) = 'Version'; $cover.Cells.Item(36, 3) = 'Date'; $cover.Cells.Item(36, 4) = 'Author'; $cover.Cells.Item(36, 5) = 'Note'
    $cover.Range('B36:E36').Style = 'FBM Header'
    $cover.Cells.Item(37, 2) = $workbookVersion
    $cover.Cells.Item(37, 3) = $asOfDate
    $cover.Cells.Item(37, 4) = $workbookAuthor
    $cover.Cells.Item(37, 5) = 'Initial release of v1.2 template (PR-A foundation + PR-B output polish).'
    $cover.Range('B37:E37').Font.Name = 'Calibri'; $cover.Range('B37:E37').Font.Size = 9

    # Legend block (rows 42-48)
    $cover.Cells.Item(42, 2) = 'Cell Style Legend'
    $cover.Cells.Item(42, 2).Style = 'FBM Subheader'
    $cover.Range('B42:D42').Merge()
    $cover.Cells.Item(43, 2) = 'Cell Style'; $cover.Cells.Item(43, 3) = 'Example'; $cover.Cells.Item(43, 4) = 'Use For'
    $cover.Range('B43:D43').Style = 'FBM Header'
    $styleNames = @('FBM Input','FBM Formula','FBM Link','FBM External','FBM Assumption')
    $examples   = @(1000, 2000, 3000, 4000, 0.10)
    $descs      = @('Hardcoded inputs (blue text)','Calculated formulas (black text)','Links from other sheets (green text)','External file links (red text)','Key assumption (yellow fill, blue bold)')
    for ($i = 0; $i -lt 5; $i++) {
        $r = 44 + $i
        $cover.Cells.Item($r, 2) = $styleNames[$i]
        $cover.Cells.Item($r, 3) = $examples[$i]
        $cover.Cells.Item($r, 3).Style = $styleNames[$i]
        if ($styleNames[$i] -eq 'FBM Assumption') { $cover.Cells.Item($r, 3).NumberFormat = $fmtPercent }
        $cover.Cells.Item($r, 4) = $descs[$i]
        $cover.Cells.Item($r, 4).Font.Name = 'Calibri'
        $cover.Cells.Item($r, 4).Font.Size = 9
    }

    # Tab color legend (rows 50-54)
    $cover.Cells.Item(50, 2) = 'Tab Color Coding'
    $cover.Cells.Item(50, 2).Style = 'FBM Subheader'
    $cover.Range('B50:D50').Merge()
    $cover.Cells.Item(51, 2) = 'Color'; $cover.Cells.Item(51, 3) = 'Purpose'; $cover.Cells.Item(51, 4) = 'Example Sheet'
    $cover.Range('B51:D51').Style = 'FBM Header'
    $cover.Cells.Item(52, 2) = 'Blue';  $cover.Cells.Item(52, 2).Interior.Color = $tabBlue;  $cover.Cells.Item(52, 2).Font.Color = $white; $cover.Cells.Item(52, 2).Font.Bold = $true
    $cover.Cells.Item(52, 3) = 'Inputs / data entry'; $cover.Cells.Item(52, 4) = 'Inputs'
    $cover.Cells.Item(53, 2) = 'Green'; $cover.Cells.Item(53, 2).Interior.Color = $tabGreen; $cover.Cells.Item(53, 2).Font.Color = $white; $cover.Cells.Item(53, 2).Font.Bold = $true
    $cover.Cells.Item(53, 3) = 'Outputs / results / summaries'; $cover.Cells.Item(53, 4) = 'Output'
    $cover.Cells.Item(54, 2) = 'Gray';  $cover.Cells.Item(54, 2).Interior.Color = $tabGray;  $cover.Cells.Item(54, 2).Font.Color = $white; $cover.Cells.Item(54, 2).Font.Bold = $true
    $cover.Cells.Item(54, 3) = 'Reference / lookup tables'; $cover.Cells.Item(54, 4) = 'Reference'

    # Confidentiality footer (row 56)
    $cover.Cells.Item(56, 2) = 'FBM Confidential — Do Not Distribute'
    $cover.Cells.Item(56, 2).Font.Italic = $true
    $cover.Cells.Item(56, 2).Font.Color = $charcoal
    $cover.Cells.Item(56, 2).Font.Size = 9
    $cover.Range('B56:E56').Merge()

    # Page setup for Cover
    $cover.PageSetup.Orientation = 1  # xlPortrait
    $cover.PageSetup.LeftMargin   = $excel.InchesToPoints(0.5)
    $cover.PageSetup.RightMargin  = $excel.InchesToPoints(0.5)
    $cover.PageSetup.TopMargin    = $excel.InchesToPoints(0.5)
    $cover.PageSetup.BottomMargin = $excel.InchesToPoints(0.5)
    $cover.PageSetup.CenterHeader = "&""Calibri,Bold""&14$workbookTitle"
    Set-FbmFooter $cover

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

    # Subtitle (FBM Subtitle style: italic charcoal, no wrap)
    $inputs.Cells.Item(3,2) = 'Blue text cells are user-editable. Do not overwrite black formula cells.'
    $inputs.Cells.Item(3,2).Style = 'FBM Subtitle'

    # Units row (centered italic charcoal)
    $inputs.Range('C4:F4').Merge()
    $inputs.Cells.Item(4,3) = "$reportingScale ($reportingCcy)"
    $inputs.Cells.Item(4,3).Style = 'FBM Units'

    # Header row
    $inputs.Range('B5:F5').Style = 'FBM Header'
    $inputs.Cells.Item(5,2)='Item'
    $inputs.Cells.Item(5,3)='FY2026'
    $inputs.Cells.Item(5,4)='FY2027'
    $inputs.Cells.Item(5,5)='FY2028'
    $inputs.Cells.Item(5,6)='FY2029'
    # Year cells: right-aligned per 2.4 convention
    $inputs.Range('C5:F5').HorizontalAlignment = -4152  # xlRight

    # Standard row heights
    Set-RowHeights $inputs 5 $true $true

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

    # Banding on data block
    Set-Banding $inputs.Range('B6:F45')

    # Named range example
    $wb.Names.Add('Revenue_Input', "=Inputs!`$C`$6:`$F`$6") | Out-Null
    $wb.Names.Add('Growth_Rate',   "=Inputs!`$C`$7:`$F`$7") | Out-Null

    # Freeze panes: one row below header_row (5) and one col right of label_col (B) -> C6.
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
    $inputs.PageSetup.LeftHeader = '&G'
    Set-FbmFooter $inputs

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

    # Units row
    $calc.Range('C4:F4').Merge()
    $calc.Cells.Item(4,3) = "$reportingScale ($reportingCcy)"
    $calc.Cells.Item(4,3).Style = 'FBM Units'

    $calc.Range('B5:F5').Style = 'FBM Header'
    $calc.Cells.Item(5,2)='Metric'
    $calc.Cells.Item(5,3)='FY2026'
    $calc.Cells.Item(5,4)='FY2027'
    $calc.Cells.Item(5,5)='FY2028'
    $calc.Cells.Item(5,6)='FY2029'
    $calc.Range('C5:F5').HorizontalAlignment = -4152  # xlRight

    Set-RowHeights $calc 5 $false $true

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

    # Banding
    Set-Banding $calc.Range('B6:F45')

    $calc.Activate()
    $calc.Range('C6').Select()
    $excel.ActiveWindow.FreezePanes = $true

    $calc.PageSetup.Orientation = 2
    $calc.PageSetup.Zoom = $false
    $calc.PageSetup.FitToPagesWide = 1
    $calc.PageSetup.FitToPagesTall = $false
    $calc.PageSetup.PrintTitleRows = '$5:$5'
    $calc.PageSetup.CenterHeader = '&"Calibri,Bold"&12Calculations'
    Set-FbmFooter $calc

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

    # KPI cards across the top (rows 4-5, columns C E G)
    $output.Cells.Item(4,3) = '=Calc!F6'
    $output.Cells.Item(4,3).Style = 'FBM KPI Big'
    $output.Cells.Item(4,3).NumberFormat = '"$"#,##0'
    $output.Cells.Item(5,3) = 'FY2029 Revenue'
    $output.Cells.Item(5,3).Style = 'FBM KPI Label'

    $output.Cells.Item(4,5) = '=Calc!F8'
    $output.Cells.Item(4,5).Style = 'FBM KPI Big'
    $output.Cells.Item(4,5).NumberFormat = '"$"#,##0'
    $output.Cells.Item(5,5) = 'FY2029 Gross Profit'
    $output.Cells.Item(5,5).Style = 'FBM KPI Label'

    $output.Cells.Item(4,7) = '=Calc!F9'
    $output.Cells.Item(4,7).Style = 'FBM KPI Big'
    $output.Cells.Item(4,7).NumberFormat = '0.0%'
    $output.Cells.Item(5,7) = 'FY2029 Gross Margin'
    $output.Cells.Item(5,7).Style = 'FBM KPI Label'

    # Units row + header row pushed to row 7/8 so KPI cards have breathing room
    $output.Rows.Item(4).RowHeight = 32
    $output.Rows.Item(5).RowHeight = 16
    $output.Rows.Item(6).RowHeight = 8  # blank breathing-room above table

    $output.Range('C7:F7').Merge()
    $output.Cells.Item(7,3) = "$reportingScale ($reportingCcy)"
    $output.Cells.Item(7,3).Style = 'FBM Units'

    $output.Range('B8:F8').Style = 'FBM Header'
    $output.Cells.Item(8,2)='KPI'
    $output.Cells.Item(8,3)='FY2026'
    $output.Cells.Item(8,4)='FY2027'
    $output.Cells.Item(8,5)='FY2028'
    $output.Cells.Item(8,6)='FY2029'
    $output.Range('C8:F8').HorizontalAlignment = -4152  # xlRight
    $output.Rows.Item(8).RowHeight = 30

    $output.Cells.Item(9,2)='Revenue'
    $output.Cells.Item(9,2).Font.Name='Calibri'; $output.Cells.Item(9,2).Font.Size=9
    $output.Range('C9:F9').Style = 'FBM Link'
    $output.Range('C9:F9').NumberFormat = $fmtCurrency
    $output.Cells.Item(9,3).Formula = '=Calc!C6'
    $output.Cells.Item(9,4).Formula = '=Calc!D6'
    $output.Cells.Item(9,5).Formula = '=Calc!E6'
    $output.Cells.Item(9,6).Formula = '=Calc!F6'

    $output.Cells.Item(10,2)='Gross Profit'
    $output.Cells.Item(10,2).Font.Name='Calibri'; $output.Cells.Item(10,2).Font.Size=9
    $output.Range('C10:F10').Style = 'FBM Link'
    $output.Range('C10:F10').NumberFormat = $fmtCurrency
    $output.Cells.Item(10,3).Formula = '=Calc!C8'
    $output.Cells.Item(10,4).Formula = '=Calc!D8'
    $output.Cells.Item(10,5).Formula = '=Calc!E8'
    $output.Cells.Item(10,6).Formula = '=Calc!F8'

    $output.Cells.Item(11,2)='Gross Margin %'
    $output.Cells.Item(11,2).Font.Name='Calibri'; $output.Cells.Item(11,2).Font.Size=9
    $output.Range('C11:F11').Style = 'FBM Link'
    $output.Range('C11:F11').NumberFormat = $fmtPercent
    $output.Cells.Item(11,3).Formula = '=Calc!C9'
    $output.Cells.Item(11,4).Formula = '=Calc!D9'
    $output.Cells.Item(11,5).Formula = '=Calc!E9'
    $output.Cells.Item(11,6).Formula = '=Calc!F9'

    # Blank breathing-room row (12) above total
    $output.Rows.Item(12).RowHeight = 8

    $output.Cells.Item(13,2) = 'Total Revenue (4Y)'
    $output.Cells.Item(13,2).Font.Bold = $true
    $output.Cells.Item(13,2).Font.Name='Calibri'; $output.Cells.Item(13,2).Font.Size=9
    $output.Cells.Item(13,3).Formula = '=SUM(C9:F9)'
    $output.Cells.Item(13,3).Style = 'FBM Total'
    $output.Cells.Item(13,3).NumberFormat = $fmtCurrency

    # Banding on data rows
    Set-Banding $output.Range('B9:F11')

    # Confidentiality footer on Output
    $output.Cells.Item(15, 2) = 'FBM Confidential — Do Not Distribute'
    $output.Cells.Item(15, 2).Font.Italic = $true
    $output.Cells.Item(15, 2).Font.Color = $charcoal
    $output.Cells.Item(15, 2).Font.Size = 9

    $output.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false  # Output: gridlines OFF
    $output.Range('C9').Select()
    $excel.ActiveWindow.FreezePanes = $true

    $output.PageSetup.Orientation = 2
    $output.PageSetup.Zoom = $false
    $output.PageSetup.FitToPagesWide = 1
    $output.PageSetup.FitToPagesTall = $false
    $output.PageSetup.PrintTitleRows = '$8:$8'
    $output.PageSetup.CenterHeader = '&"Calibri,Bold"&12Output Summary'
    Set-FbmFooter $output

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

    # Reference data — expanded with full named-range set
    $refData = @(
        @('Discount rate',    0.10,   'Discount_Rate',    $fmtPercent),
        @('Tax rate',         0.21,   'Tax_Rate',         $fmtPercent),
        @('WACC',             0.085,  'WACC',             $fmtPercent),
        @('Inflation rate',   0.025,  'Inflation_Rate',   $fmtPercent),
        @('FY start month',   1,      'FY_Start',         $null),
        @('FY end month',     12,     'FY_End',           $null),
        @('Working days/yr',  252,    $null,              $null)
    )
    for ($i=0; $i -lt $refData.Count; $i++) {
        $r = 6 + $i
        $ref.Cells.Item($r,2) = $refData[$i][0]
        $ref.Cells.Item($r,2).Font.Name='Calibri'; $ref.Cells.Item($r,2).Font.Size=9
        $ref.Cells.Item($r,3) = $refData[$i][1]
        $ref.Cells.Item($r,3).Style = 'FBM Input'
        if ($refData[$i][3]) { $ref.Cells.Item($r,3).NumberFormat = $refData[$i][3] }
        if ($refData[$i][2]) {
            $wb.Names.Add($refData[$i][2], "=Reference!`$C`$$r") | Out-Null
        }
    }

    # Banding
    Set-Banding $ref.Range('B6:C30')

    $ref.Activate()
    $ref.Range('C6').Select()
    $excel.ActiveWindow.FreezePanes = $true

    $ref.PageSetup.Orientation = 1
    $ref.PageSetup.CenterHeader = '&"Calibri,Bold"&12Reference'
    Set-FbmFooter $ref

    # ============================================
    # WORKBOOK PROPERTIES (File -> Info -> Properties)
    # ============================================
    $wb.BuiltinDocumentProperties.Item('Title')    = $workbookTitle
    $wb.BuiltinDocumentProperties.Item('Subject')  = $workbookSubject
    $wb.BuiltinDocumentProperties.Item('Author')   = $workbookAuthor
    $wb.BuiltinDocumentProperties.Item('Company')  = $workbookCompany
    $wb.BuiltinDocumentProperties.Item('Category') = $workbookCategory
    try { $wb.BuiltinDocumentProperties.Item('Last Author') = $workbookAuthor } catch {}

    # Make Cover the active sheet on open and hide gridlines on Cover
    $cover.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false

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
