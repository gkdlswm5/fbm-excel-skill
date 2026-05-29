# FBM Excel Formatting Standard

**Version 1.0** &middot; Owner: Andrew Kim, Operations FP&A &middot; Last updated: 2026-05-21

This document defines the visual and structural standards for all FBM Excel workbooks. The companion file `FBM-Excel-Template.xlsx` implements every rule below as pre-built cell styles, named ranges, and example sheets — copy it to start any new workbook.

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
| Light grey | ![lt](https://placehold.co/14x14/E7E6E6/E7E6E6.png) | `#E7E6E6` | Subheader fill, banding |

### 1.2 Fonts

- **Headers / titles:** Rockwell (per FBM brand). Falls back to Calibri Bold if Rockwell unavailable.
- **Body:** Calibri 9
- **Bold headers:** Calibri 11 bold

### 1.3 Logo

- File: `C:\ak\fbm-standards\fbm-logo.jpeg` (extracted from official source)
- Placement: Cover sheet, top-left, ~200&times;90 pt
- Print: include in left header (`&G`) on print-facing sheets

---

## 2. Structure

### 2.1 Sheet order & tab colors

| # | Sheet | Tab color | Purpose |
|---|---|---|---|
| 1 | Cover | Navy `#093254` | Title, metadata, legend |
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
- **Row 3:** optional subtitle / instruction (italic, charcoal)
- **Row 5:** column headers (FBM Header style)
- **Row 6+:** data
- **Freeze panes:** at C6 (header row + label column locked)
- **AutoFilter:** applied to header row on Inputs and tabular sheets

---

## 3. Data formatting

### 3.1 Number formats (single standard each)

| Type | Format | Example |
|---|---|---|
| Currency | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)` | $1,234 / $(1,234) / - |
| Currency (cents) | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` | $1,234.56 |
| Number | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` | 1,234 / (1,234) / - |
| Percent | `0.0%;(0.0%);"-"` | 12.3% / (12.3%) / - |
| Multiple | `0.0"x"` | 4.5x |
| Date | `mm/dd/yyyy` | 05/21/2026 |
| Year | text (`@`), centered | 2026 (not 2,026) |

### 3.2 Negative numbers

**Parentheses, not minus signs.** All standard number formats above include `(...)` for negatives. Red text is reserved for external links only — don't use red for negative numbers.

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

Apply via icon sets, color scales, or direct conditional formats. The template's Calc sheet shows an icon-set example on Gross Margin %.

### 3.6 Pre-built cell styles in the template

| Style name | Purpose |
|---|---|
| FBM Title | 14pt Rockwell, navy — sheet titles (row 2) |
| FBM Header | Navy fill, white bold, centered — header row (row 5) |
| FBM Subheader | Light grey fill, navy bold — section dividers |
| FBM Input | Blue text, comma number format |
| FBM Input $ | Blue text, currency format |
| FBM Input % | Blue text, percent format |
| FBM Formula | Black text, comma number format |
| FBM Formula $ | Black text, currency format |
| FBM Formula % | Black text, percent format |
| FBM Link | Green text, comma number format |
| FBM External | Red text, comma number format |
| FBM Assumption | Yellow fill, blue bold — key assumption |
| FBM Total | Top single + bottom double border, bold |
| FBM Date | Black text, mm/dd/yyyy |
| FBM Year | Black text, centered, text format |

---

## 4. Rules & conventions

### 4.1 Formulas

- **No hardcoded values inside formulas.** Use cell references. `=B5*(1+$C$5)` not `=B5*1.05`.
- **All assumptions belong in `Inputs` or `Reference`,** never inline in `Calc` or `Output`.
- **Use named ranges** for any constant referenced more than twice (`Tax_Rate`, `Discount_Rate`, etc.). The template defines `Revenue_Input`, `Growth_Rate`, `Tax_Rate`, `Discount_Rate` as starter examples.
- **Wrap divisions in `IFERROR(...)`** to prevent `#DIV/0!`.
- **No formula errors permitted.** Before sharing, verify zero `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`.
- **Document hardcoded source values** in a cell comment or adjacent cell:
  Format: `Source: [System/Document], [Date], [Reference/URL]`

### 4.2 Print setup

| Setting | Value |
|---|---|
| Orientation | Landscape (data sheets), Portrait (Cover, Reference) |
| Scaling | Fit to 1 page wide, unlimited tall |
| Margins | Left/Right 0.5", Top 0.75", Bottom 0.5" |
| Header | Center: sheet title (Calibri Bold 12); Left: logo (`&G`) |
| Footer | Left: `&F` (file path); Center: `&D` (date); Right: `&P of &N` (page) |
| Repeat rows | `$5:$5` (header row on every page) |

### 4.3 File naming convention

Format: `FBM - [Subject] - [YYYY.MM.DD].xlsx`

Examples:
- `FBM - REC Leases - 2026.05.20.xlsx`
- `FBM - Leadership Update - 2026.04.29.xlsx`
- `FBM - Q2 Forecast v3 - 2026.05.15.xlsx`

Version suffixes: `v1`, `v2`, ... or initials (`ak`) for working copies. Date in filename always reflects the data as-of date, not the save date.

### 4.4 Workbook hygiene

- Cover sheet metadata filled in (author, date, version, purpose)
- No leftover scratch sheets named `Sheet1`, `Test`, etc.
- Hide (don't delete) work-in-progress sheets if needed for audit
- Set `Cover` as the active sheet before saving
- Clear filters before saving final version
- Remove any external file links not intended to persist (`Data → Edit Links`)

---

## 5. Quick checklist before sharing

- [ ] Cover sheet metadata complete
- [ ] All tab colors set per convention
- [ ] No `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?` errors
- [ ] No hardcoded values inside formulas
- [ ] Inputs are blue, formulas are black, cross-sheet links are green
- [ ] Number formats applied (zeros as `-`, negatives in parens)
- [ ] Freeze panes on data sheets at C6
- [ ] AutoFilter applied to header rows
- [ ] Print setup: landscape + fit-to-page-wide + repeat row 5
- [ ] Filename follows `FBM - [Subject] - [YYYY.MM.DD].xlsx`
- [ ] Cover sheet active before saving

---

## Appendix A — Working with Claude

To have Claude apply this standard to any Excel file:

1. Standards are auto-loaded from `C:\ak\fbm-standards\` (referenced in CLAUDE.md and Claude's memory).
2. Just say "apply the FBM standard to this file" or "reformat to FBM standard" — Claude will pick up the colors, fonts, number formats, tab conventions, and cell styles from this document.
3. To update standards: edit this file + `FBM-Excel-Template.xlsx`. Claude will read the latest version on each session.
