# FBM Excel Formatting Standard

**Version 1.2** &middot; Owner: Andrew Kim, Operations FP&A &middot; Last updated: 2026-05-29

This document defines the visual and structural standards for all FBM Excel workbooks. The companion file `FBM-Excel-Template.xlsx` implements every rule below as pre-built cell styles, named ranges, and example sheets — copy it to start any new workbook.

Changelog:
- **v1.2 (2026-05-29)** &mdash; PR-A foundation + PR-B output polish: color role lock, units row, row-height hygiene, sub-item indent rule, header alignment, banding, gridline rule, variance formats, edge-case formats, KPI styles, expanded Cover sheet content, hyperlinked TOC, footer metadata, workbook properties.
- **v1.1 (2026-05-29)** &mdash; subtitle no-wrap rule, dynamic freeze panes, sign convention (&sect;4.2), pivot refresh rule (&sect;4.3).
- **v1.0 (2026-05-21)** &mdash; initial standard.

---

## 1. Visual identity

### 1.1 Brand colors (from FBM PPTX theme)

| Role | Color | Hex | Use for |
|---|---|---|---|
| Primary navy | ![navy](https://placehold.co/14x14/093254/093254.png) | `#093254` | Headers, titles, tab color for Cover |
| Burgundy | ![burgundy](https://placehold.co/14x14/6A1831/6A1831.png) | `#6A1831` | Accent, callouts |
| Slate blue | ![slate](https://placehold.co/14x14/4E7087/4E7087.png) | `#4E7087` | Secondary headers |
| Charcoal | ![charcoal](https://placehold.co/14x14/636466/636466.png) | `#636466` | Body text, muted labels |
| Forest green | ![forest](https://placehold.co/14x14/005E34/005E34.png) | `#005E34` | Output tab color, positive trend |
| Sage | ![sage](https://placehold.co/14x14/708573/708573.png) | `#708573` | Soft accent |
| Light grey | ![lt](https://placehold.co/14x14/E7E6E6/E7E6E6.png) | `#E7E6E6` | Subheader fill |
| Banding grey | ![band](https://placehold.co/14x14/F7F7F7/F7F7F7.png) | `#F7F7F7` | Alternating data-row fill |

### 1.1.1 Color role lock

Each color has a single, locked role. Don't mix roles, don't introduce ad-hoc colors.

| Color | Locked role |
|---|---|
| Navy | Primary headers, sheet titles, Cover tab color, KPI big numbers |
| Slate blue | Secondary / grouping headers within a section |
| Burgundy | Callouts only &mdash; warnings, footnotes, audit issues |
| Sage | Soft category accents &mdash; chart fills, sub-totals |
| Forest green | Output tab color; "shipped" / "approved" status text |
| Charcoal | Body text, muted labels, soft borders, subtitle text |
| Light grey | Subheader fill, table-section dividers |
| Banding grey | Alternating data-row fill on tabular sheets only |

Reviewer rule: if a cell uses a color outside this matrix, that cell will be flagged.

### 1.2 Fonts

- **Headers / titles:** Rockwell (per FBM brand). Falls back to Calibri Bold if Rockwell unavailable.
- **Body:** Calibri 9
- **Bold headers:** Calibri 11 bold
- **KPI big number:** Rockwell 24, navy, bold
- **KPI label:** Calibri 9, charcoal, italic

### 1.3 Logo

- File: `C:\ak\fbm-standards\fbm-logo.jpeg` (extracted from official source)
- Placement: Cover sheet, top-left, ~200&times;90 pt
- Print: include in left header (`&G`) on print-facing sheets

---

## 2. Structure

### 2.1 Sheet order & tab colors

| # | Sheet | Tab color | Purpose |
|---|---|---|---|
| 1 | Cover | Navy `#093254` | Title, metadata, legend, TOC |
| 2 | Inputs | Blue `#1F4E79` | User-editable data entry |
| 3 | Calc | (none) | Calculations / intermediate |
| 4 | Output | Green `#005E34` | KPIs, summaries, deliverable views |
| 5 | Reference | Gray `#A5A5A5` | Lookup tables, constants, named ranges |

### 2.2 Layout conventions

- **Column A:** narrow gutter, width = 5
- **Column B:** labels, width = 30
- **Columns C onward:** data, default width = 15 (multiples of 5 — use 5, 10, 15, 20, 25)
- **Row 1:** blank (top margin)
- **Row 2:** sheet title (FBM Title style)
- **Row 3:** optional subtitle / instruction (italic, charcoal). **Do not enable `Wrap Text` on the subtitle cell** — leave it single-line, left-aligned, and let long text overflow into the empty cells to the right. Wrapping a long sentence inside a 30-width label column balloons the row height and breaks the layout. If the message must wrap, merge the subtitle cell across the data columns (e.g. `B3:H3`) so it has horizontal room before turning wrap on.
- **Row 4:** optional **units row** (FBM Units style: italic charcoal, centered, no wrap). Use literal labels: `$ thousands`, `$ millions`, `%`, `bps`, `units`, `# heads`, `days`. Eliminates the most common reviewer question. Skip the row entirely if the sheet doesn't have a single uniform unit.
- **Row 5:** column headers (FBM Header style) — this is the default position, but the header row may sit lower (e.g. row 6 or 7) when extra subtitle/instruction rows are needed.
- **Row 6+:** data (one row below the header row)
- **Freeze panes:** computed dynamically from the actual layout — freeze one row *below* the header row and one column *right* of the rightmost label column. With the default layout (header in row 5, labels in column B only) this resolves to `C6`. If the header sits in row 7 and labels span `B:C`, freeze at `D8`. Never hardcode `C6` when the header is elsewhere.
- **AutoFilter:** applied to the header row on Inputs and tabular sheets.
- **Banding:** tabular sheets get a `#F7F7F7` fill on alternating data rows (e.g. rows 6, 8, 10, &hellip;). Apply via conditional format `=MOD(ROW(),2)=0` or via `apply_styles.apply_banding(ws, range)`.
- **Sub-line-item indent:** indent via `Alignment(indent=N)`, **never** leading spaces. Sort-safe, copy-safe, and machine-readable.
- **Decimal-place consistency:** every cell in the same column uses the same precision. Never mix `$1,234` and `$1,234.56` in one column &mdash; either both round, or both show cents.

### 2.3 Row-height hygiene

Locked row heights so every sheet looks identical:

| Row | Purpose | Height (pt) |
|---|---|---|
| 2 | Title | 24 |
| 3 | Subtitle (when used) | 16 |
| 4 | Units (when used) | 16 |
| 5 | Header | 30 |
| 6+ | Data | 16 |
| Row above totals | Blank breathing-room | 8 |

Helper: `apply_styles.apply_row_heights(ws, header_row=5, has_subtitle=True, has_units=False)`.

### 2.4 Header / label alignment

| Cell role | Alignment |
|---|---|
| Label column (B) | Left, vertically centered |
| Numeric headers (e.g. `FY2026`, `Q1`) | Right or center |
| Text headers (e.g. `Region`, `Item`) | Center |
| Numeric data | Right |
| Text data | Left |
| Units row | Center, italic charcoal |
| KPI big number | Center |

### 2.5 Gridlines

- **Cover sheet:** gridlines **off**.
- **Output sheet:** gridlines **off** (it's the deliverable view).
- **Inputs, Calc, Reference:** gridlines **on** (analyst working sheets).

Set via `ws.sheet_view.showGridLines = False` (openpyxl) or `$ws.DisplayGridlines = $false` (COM). Helper: `apply_styles.hide_gridlines(ws)`.

---

## 3. Data formatting

### 3.1 Number formats

| Type | Format | Example |
|---|---|---|
| Currency | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)` | $1,234 / $(1,234) / - |
| Currency (cents) | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` | $1,234.56 |
| Number | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` | 1,234 / (1,234) / - |
| Percent | `0.0%;(0.0%);"-"` | 12.3% / (12.3%) / - |
| Multiple | `0.0"x"` | 4.5x |
| Date | `mm/dd/yyyy` | 05/21/2026 |
| Year | text (`@`), centered | 2026 (not 2,026) |
| **Variance ($)** | `+$#,##0;($#,##0);"-"` | +$1,234 / ($1,234) / - |
| **Variance (number)** | `+#,##0;(#,##0);"-"` | +1,234 / (1,234) / - |
| **Variance (%)** | `+0.0%;(0.0%);"-"` | +1.2% / (1.2%) / - |
| **Variance (bps)** | `+0" bps";(0)" bps";"-"` | +12 bps / (12) bps / - |
| Ratio | `0.00` | 1.25 |
| Basis points | `0" bps"` | 125 bps |
| Share count | `#,##0` | 12,345,678 |
| Share price | `$#,##0.00` | $123.45 |

`n/a` cells: use the literal string `n/a`. Not blank, not zero, not `#N/A`.

### 3.2 Negative numbers

**Parentheses, not minus signs.** All standard number formats above include `(...)` for negatives. Red text is reserved for external links only — don't use red for negative numbers.

**Cost / expense display convention:** costs are displayed as positive numbers (they're "uses of cash" in their own column). Only true contra-revenue items &mdash; rebates, discounts, returns, allowances &mdash; are shown as negative. This keeps the income-statement view clean and lets variance columns be the only place a number's sign carries meaning.

Note: this section is about *display*. The rule about whether a value should *be* negative — the "positive good, negative bad" convention for variance formulas — lives in &sect;4.2.

### 3.3 Zeros

Display zeros as `-` (dash) via the number format. Always include `;"-"` in the third format section. This applies to percentages too.

### 3.4 Cell text colors (font color convention)

| Text color | Hex | Meaning |
|---|---|---|
| Blue | `#0000FF` | Hardcoded inputs / assumptions user will change |
| Black | `#000000` | Formulas and calculations |
| Green | `#008000` | Links to other sheets in same workbook |
| Red | `#FF0000` | External file links |
| **Yellow fill + bold blue** | bg `#FFFF00`, fg `#0000FF` | Key assumption requiring attention |

### 3.5 Conditional formatting palette (good / neutral / bad)

- **Good:** font `#006100`, fill `#C6EFCE`
- **Neutral:** font `#9C5700`, fill `#FFEB9C`
- **Bad:** font `#9C0006`, fill `#FFC7CE`

Rules of engagement:
- **Variance %, growth %, margin % &mdash; always** get the good/neutral/bad palette. The &sect;4.2 sign convention guarantees green=good without metric-specific logic.
- **Raw $ &mdash; never** get CF. Only variance from $ does.
- **Headcount &mdash; never** CF (context-dependent &mdash; under or over plan can both be problems).

Apply via icon sets, color scales, or direct conditional formats. The template's Calc sheet shows an icon-set example on Gross Margin %.

### 3.6 Pre-built cell styles in the template

| Style name | Purpose |
|---|---|
| FBM Title | 14pt Rockwell, navy &mdash; sheet titles (row 2) |
| FBM Subtitle | 9pt Calibri italic charcoal, no wrap &mdash; subtitle (row 3) |
| FBM Units | 9pt Calibri italic charcoal, centered &mdash; units row (row 4) |
| FBM Header | Navy fill, white bold, centered &mdash; header row |
| FBM Subheader | Light grey fill, navy bold &mdash; section dividers |
| FBM Input | Blue text, comma number format |
| FBM Input $ | Blue text, currency format |
| FBM Input % | Blue text, percent format |
| FBM Formula | Black text, comma number format |
| FBM Formula $ | Black text, currency format |
| FBM Formula % | Black text, percent format |
| FBM Link | Green text, comma number format |
| FBM External | Red text, comma number format |
| FBM Assumption | Yellow fill, blue bold &mdash; key assumption |
| FBM Total | Top single + bottom double border, bold |
| FBM Date | Black text, mm/dd/yyyy |
| FBM Year | Black text, centered, text format |
| **FBM Variance** | Black text, forced-sign number format |
| **FBM Variance $** | Black text, forced-sign currency format |
| **FBM Variance %** | Black text, forced-sign percent format |
| **FBM Variance bps** | Black text, forced-sign bps format |
| **FBM Band** | Banding grey fill (`#F7F7F7`) for alternating data rows |
| **FBM KPI Big** | 24pt Rockwell bold navy, centered |
| **FBM KPI Label** | 9pt Calibri italic charcoal, centered |

### 3.7 Total rows

- Style: `FBM Total` (single thin top border + double bottom border, bold).
- **Always include an 8pt blank row above the total** for visual breathing room. The thin top border on the Total style + the blank row above is what makes the total read as a total at a glance.
- Sub-totals use slate-blue fill + bold instead of the double bottom border.

---

## 4. Rules & conventions

### 4.1 Formulas

- **No hardcoded values inside formulas.** Use cell references. `=B5*(1+$C$5)` not `=B5*1.05`.
- **All assumptions belong in `Inputs` or `Reference`,** never inline in `Calc` or `Output`.
- **Use named ranges** for any constant referenced more than twice. Standard named ranges defined in the template: `AsOfDate`, `Scale`, `ReportingCurrency`, `FY_Start`, `FY_End`, `Discount_Rate`, `Tax_Rate`, `WACC`, `Inflation_Rate`, `Revenue_Input`, `Growth_Rate`.
- **Wrap divisions in `IFERROR(...)`** to prevent `#DIV/0!`.
- **No formula errors permitted.** Before sharing, verify zero `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`.
- **Variance columns follow &sect;4.2 sign convention** &mdash; subtraction direction is metric-dependent.
- **Document hardcoded source values** in a cell comment or adjacent cell:
  Format: `Source: [System/Document], [Date], [Reference/URL]`

### 4.2 Sign convention &mdash; "positive good, negative bad"

FBM rule: in every variance, delta, change, or "vs" column, the formula MUST be constructed so a **positive number is a favorable outcome** and a **negative number is unfavorable**. The reader should never have to think "is + good or bad here?" &mdash; it is always good.

The direction of subtraction depends on whether *higher* or *lower* is favorable for that metric.

| Metric | Higher is | Variance formula | Example: `+$1M` means |
|---|---|---|---|
| Revenue, gross profit, EBITDA, cash | better | `Actual - Plan` | Beat plan by $1M (good) |
| COGS, OpEx, SG&A, interest expense | worse | `Plan - Actual` | Spent $1M less than plan (good) |
| Headcount (cost view) | worse | `Plan - Actual` | Under plan by 1 head (good) |
| DSO, DPO, inventory days, cycle time | worse | `Prior - Current` | Improved by 1 day (good) |
| Margin %, ROIC, conversion % | better | `Actual - Plan` | Above plan by N bps (good) |
| Churn %, defect rate, returns % | worse | `Plan - Actual` | Below plan by N bps (good) |

This convention pairs with:

- **Number format** &mdash; use the forced-sign variance formats from &sect;3.1 (`+#,##0;(#,##0);"-"`, `+0.0%;(0.0%);"-"`, `+0" bps";(0)" bps";"-"`).
- **Conditional formatting** (&sect;3.5) &mdash; green = positive = good, red = negative = bad. With the sign convention enforced, the good/neutral/bad palette can be applied to any variance column without metric-specific logic.
- **Column header** &mdash; name variance columns `Fav/(Unfav) vs Plan`, `Fav/(Unfav) YoY`, or `Δ Fav/(Unfav)` so the convention is explicit to the reader.

For non-obvious cases (cost variances especially), document the subtraction direction in column B or as a cell comment, e.g. `COGS variance = Plan - Actual (under = favorable)`.

### 4.3 Pivot tables

Pivot tables must survive a refresh without losing the visual layout the analyst set up.

- **Disable "Autofit column widths on update"** so column widths set per section 2.2 (label col = 30, data cols = 15) aren't blown away on every refresh.
- **Enable "Preserve cell formatting on update"** so manual fonts, borders, and number formats stick across refreshes.
- Apply these to every pivot in the workbook, not just the first one.

Set via:

| Surface | Setting |
|---|---|
| Excel UI | PivotTable Analyze &rarr; Options &rarr; Layout & Format &rarr; uncheck *Autofit column widths on update*, check *Preserve cell formatting on update* |
| Excel COM (PowerShell / VBA) | `$pt.HasAutoFormat = $false; $pt.PreserveFormatting = $true` |
| openpyxl | `pivot.useAutoFormatting = False; pivot.preserveFormatting = True` (helper: `apply_styles.disable_pivot_autofit(wb)`) |

### 4.4 Print setup

| Setting | Value |
|---|---|
| Orientation | Landscape (data sheets), Portrait (Cover, Reference) |
| Scaling | Fit to 1 page wide, unlimited tall; scale floor **60%** &mdash; reorganize before shrinking further |
| Margins | Left/Right 0.5", Top 0.75", Bottom 0.5" |
| Header | Center: sheet title (Calibri Bold 12); Left: logo (`&G`) |
| Footer | Left: `&F` (file path); Center: `&D` &middot; `v[Version]` &middot; `As of [AsOfDate]`; Right: `&P of &N` |
| Repeat rows | Header row (default `$5:$5`, adjust if header sits elsewhere) |
| Confidentiality footer | Right side or center bottom: `FBM Confidential — Do Not Distribute` on Cover and Output |

### 4.5 Cover sheet content

The Cover is the first and most-photographed page. It must include:

- **Logo** (top-left, ~200x90 pt)
- **Workbook title** (FBM Title, 20pt)
- **Subtitle** (e.g. "Q2 Forecast &mdash; April Refresh")
- **Metadata block** (each on its own row, label bold Calibri 10):
  - Prepared by
  - Department
  - **As-of date** (named range: `AsOfDate`)
  - **Reporting scale** (named range: `Scale`; one of `$ thousands` / `$ millions` / `actuals`)
  - **Reporting currency** (named range: `ReportingCurrency`; default `USD`)
  - Contact email
  - Version
  - File location
  - Purpose
- **Hyperlinked TOC** (one row per sheet, `FBM Link` style):
  - `=HYPERLINK("#Inputs!A1","Inputs")`
  - `=HYPERLINK("#Calc!A1","Calc")`
  - `=HYPERLINK("#Output!A1","Output")`
  - `=HYPERLINK("#Reference!A1","Reference")`
- **Data sources block** (3-5 rows: source system, last refreshed, owner)
- **Version history mini-table** (5 most recent rows: version, date, author, note)
- **Legend** (cell styles + tab color coding, as in template v1.0)
- **Confidentiality footer** at the bottom of the printable area: italic charcoal, `FBM Confidential — Do Not Distribute`

### 4.6 File naming convention

Format: `FBM - [Subject] - [YYYY.MM.DD].xlsx`

Examples:
- `FBM - REC Leases - 2026.05.20.xlsx`
- `FBM - Leadership Update - 2026.04.29.xlsx`
- `FBM - Q2 Forecast v3 - 2026.05.15.xlsx`

Version suffixes: `v1`, `v2`, ... or initials (`ak`) for working copies. Date in filename always reflects the data as-of date, not the save date.

### 4.7 Workbook hygiene

- Cover sheet metadata filled in (author, date, version, purpose, scale, currency, contact)
- **Workbook properties** filled in (File &rarr; Info &rarr; Properties):
  - Title = workbook title
  - Subject = purpose
  - Author = preparer
  - Company = `Foundation Building Materials`
  - Category = department / use-case (e.g. `Operations FP&A — Leadership`)
- No leftover scratch sheets named `Sheet1`, `Test`, `Sheet 1 (2)`, etc.
- Hide (don't delete) work-in-progress sheets if needed for audit
- Set `Cover` as the active sheet before saving
- Clear filters before saving final version
- Remove any external file links not intended to persist (`Data → Edit Links`)

---

## 5. Quick checklist before sharing

- [ ] Cover sheet metadata complete (incl. as-of date, scale, currency, contact)
- [ ] Workbook properties set (Title, Subject, Author, Company, Category)
- [ ] Hyperlinked TOC on Cover; data sources + version history blocks populated
- [ ] All tab colors set per convention
- [ ] Gridlines off on Cover and Output; on for Inputs/Calc/Reference
- [ ] Banding applied on tabular sheets
- [ ] Row 4 units row populated where the sheet has a uniform unit
- [ ] Standard row heights applied (title 24, subtitle/units 16, header 30, data 16, blank-above-total 8)
- [ ] Sub-items indented via `Alignment(indent=N)`, no leading spaces
- [ ] Decimal places consistent within each column
- [ ] No `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?` errors
- [ ] No hardcoded values inside formulas
- [ ] Inputs are blue, formulas are black, cross-sheet links are green
- [ ] Number formats applied (zeros as `-`, negatives in parens)
- [ ] Variance columns use forced-sign formats AND follow "positive good, negative bad" sign convention
- [ ] CF good/neutral/bad palette applied to variance %, growth %, margin %; never to raw $
- [ ] Freeze panes on data sheets one row below header and one column right of last label column (default `C6`)
- [ ] Subtitle / instruction rows have `Wrap Text` off (overflow right) — no abnormally tall rows
- [ ] All pivot tables have "Autofit column widths on update" OFF and "Preserve cell formatting on update" ON
- [ ] AutoFilter applied to header rows
- [ ] Print setup: landscape + fit-to-page-wide (no scale below 60%) + repeat header row
- [ ] Print footer includes file path, date, version, as-of date, page number
- [ ] Confidentiality footer present on Cover and Output
- [ ] Filename follows `FBM - [Subject] - [YYYY.MM.DD].xlsx`
- [ ] Cover sheet active before saving

---

## Appendix A — Working with Claude

To have Claude apply this standard to any Excel file:

1. Standards are auto-loaded from `C:\ak\fbm-standards\` (referenced in CLAUDE.md and Claude's memory).
2. Just say "apply the FBM standard to this file" or "reformat to FBM standard" — Claude will pick up the colors, fonts, number formats, tab conventions, and cell styles from this document.
3. To update standards: edit this file + `FBM-Excel-Template.xlsx`. Claude will read the latest version on each session.

## Appendix B — Future work

See `references/roadmap.md` for the polish backlog. PR-A (foundation) and PR-B (output polish) shipped in v1.2. PR-C (rigor &mdash; data validation, sheet protection, edge-case formats, print preview discipline) remains queued.
