---
name: fbm-excel
description: Apply Foundation Building Materials (FBM) Excel formatting standards — brand colors (navy/burgundy/forest), Rockwell/Calibri fonts, cell styles, number formats, tab colors (blue=input, green=output, gray=reference), named ranges, print setup, file naming. Use whenever the user works with .xlsx files in an FBM context, mentions FBM FP&A workbooks, asks to "apply the FBM standard", "reformat to FBM", "standardize this Excel", or wants Excel formatting to match the FBM brand. Trigger when starting a new FBM workbook, reformatting an existing one (FP&A reviews, leadership updates, REC leases, commission reports), or when the user shares an .xlsx that should look "on brand". Even if the user just says "format this spreadsheet" without explicitly saying "FBM", trigger if there's any Foundation Building Materials context in the conversation, file path, or folder.
---

# FBM Excel Formatting Standard

This skill applies the Foundation Building Materials (FBM) Excel formatting standard to workbooks. The user is in Operations FP&A at FBM and uses this for all FP&A deliverables (leadership reviews, leases, forecasts, commission reports).

## Bundled resources

This skill is **self-contained** — all resources are bundled inside the skill folder. Use skill-relative paths.

| Path | Purpose |
|---|---|
| `references/standards.md` | Full standard: colors, fonts, number formats, conventions, checklist |
| `references/roadmap.md` | Polish backlog (PR-A / PR-B / PR-C waves) &mdash; future-work, not yet shipped |
| `assets/template.xlsx` | Working template with 25 pre-built cell styles, named ranges, example sheets (Cover, Inputs, Calc, Output, Reference). **Note:** v1.2 binary not yet regenerated &mdash; run `scripts/build-template.ps1` on Windows to refresh. |
| `assets/fbm-logo.jpeg` | Official FBM logo for headers / cover sheets |
| `scripts/build-template.ps1` | Source script that generated the template (Windows + Excel COM, for reference) |
| `scripts/apply_styles.py` | Cross-platform openpyxl helper to inject FBM styles into any workbook |

**Always read `references/standards.md` first** when this skill triggers — it's the authoritative rule set.

## Workflow

### Step 1 — Read the standard
Read `references/standards.md` for the complete rule set (colors, fonts, formats, conventions, checklist).

### Step 2 — Choose your approach
Three ways to apply, in order of preference:

1. **Start from the bundled template** (best for new workbooks): copy `assets/template.xlsx` to the target location, rename per the file-naming convention (`FBM - [Subject] - [YYYY.MM.DD].xlsx`), fill in data.

2. **Inject styles into an existing workbook** (best for reformatting): use `scripts/apply_styles.py` to add the 15 FBM-prefixed named styles to a target file. Then apply them to cells per the conventions.

3. **Format manually** (best for one-off cells): use the brand color hex codes, fonts, and number-format strings directly from the standard.

### Step 3 — Preserve existing conventions
If the user is editing an existing FBM file with its own established format, **preserve those conventions**. Only apply this standard for:
- Brand-new workbooks
- Net-new sheets in an existing workbook
- Explicit requests to "reformat to FBM standard"

### Step 4 — Verify before delivering
Run the checklist in section 5 of `references/standards.md`. Non-negotiables: zero formula errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`), no hardcoded values inside formulas, correct number formats, correct cell-text color convention.

## Quick reference (canonical version in `references/standards.md`)

### Brand colors
| Role | Hex |
|---|---|
| Navy (primary) | `#093254` |
| Burgundy | `#6A1831` |
| Slate blue | `#4E7087` |
| Charcoal | `#636466` |
| Forest green | `#005E34` |
| Sage | `#708573` |
| Light grey | `#E7E6E6` |
| Banding grey | `#F7F7F7` |

Colors have **locked roles** &mdash; see &sect;1.1.1 of `standards.md`. Don't introduce ad-hoc colors.

### Fonts
- Headers/titles: **Rockwell** (Calibri Bold fallback)
- Body: **Calibri 9**
- Bold headers: **Calibri 11 bold**

### Tab colors
Blue → Inputs · Green → Outputs · Gray → Reference · Navy → Cover

### Cell text color convention
- **Blue** text = hardcoded input
- **Black** text = formula
- **Green** text = cross-sheet link
- **Red** text = external file link
- **Yellow fill + bold blue** = key assumption

### Number formats
| Type | Format |
|---|---|
| Currency | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)` |
| Number | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)` |
| Percent | `0.0%;(0.0%);"-"` |
| Multiple | `0.0"x"` |
| Date | `mm/dd/yyyy` |
| Year | text `@`, centered |
| Variance ($) | `+$#,##0;($#,##0);"-"` |
| Variance (#) | `+#,##0;(#,##0);"-"` |
| Variance (%) | `+0.0%;(0.0%);"-"` |
| Variance (bps) | `+0" bps";(0)" bps";"-"` |
| Ratio | `0.00` |
| Bps | `0" bps"` |
| Share price | `$#,##0.00` |

Negatives in parentheses (not minus). Zeros shown as `-`. `n/a` cells use the literal string `n/a` (not blank, not zero).

### Sign convention — "positive good, negative bad"
Variance / delta columns: construct the formula so favorable = positive, unfavorable = negative.
- Revenue, profit, margin: `Actual - Plan`
- Cost, OpEx, headcount, DSO, churn: `Plan - Actual` (or `Prior - Current` for time-based)
- Header columns: name them `Fav/(Unfav) vs Plan`. See §4.2 of `references/standards.md`.

### Layout
- Col A = 5 (gutter), Col B = 30 (labels), other cols = 15 default (all multiples of 5)
- Row 2 = title (24pt), row 3 = optional subtitle (**wrap OFF — overflow right**), row 4 = optional units row (italic charcoal centered: `$ thousands` / `%` / `bps` / `units`), row 5 = headers (30pt), data rows = 16pt, autofilter on header row
- Locked row heights: title 24, subtitle/units 16, header 30, data 16, blank-above-total 8 pt
- Sub-items indent via `Alignment(indent=N)`, never leading spaces
- Banding: `#F7F7F7` fill on alternating data rows via CF `=MOD(ROW(),2)=0`
- Gridlines OFF on Cover and Output; ON for Inputs/Calc/Reference
- Header / label alignment: labels left, numeric headers right, text headers center
- Freeze panes: dynamic — one row below the header row, one col right of the rightmost label column (default `C6`; never hardcode when the header isn't on row 5)
- Sheet order: Cover → Inputs → Calc → Output → Reference

### Pivot tables
- "Autofit column widths on update" **OFF**, "Preserve cell formatting on update" **ON** — applies to every pivot in the workbook.
- COM: `$pt.HasAutoFormat=$false; $pt.PreserveFormatting=$true`
- openpyxl: `apply_styles.disable_pivot_autofit(wb)`

### File naming
`FBM - [Subject] - [YYYY.MM.DD].xlsx`

## Cell styles bundled in `assets/template.xlsx`

| Style | Purpose |
|---|---|
| FBM Title | 14pt Rockwell, navy &mdash; sheet titles |
| FBM Subtitle | 9pt italic charcoal, no wrap &mdash; row 3 |
| FBM Units | 9pt italic charcoal, centered &mdash; row 4 (`$ thousands`, `%`, `bps`, etc.) |
| FBM Header | Navy fill, white bold, centered &mdash; header row |
| FBM Subheader | Light grey fill, navy bold |
| FBM Input / Input $ / Input % | Blue text &mdash; hardcoded inputs |
| FBM Formula / Formula $ / Formula % | Black text &mdash; calculated values |
| FBM Variance / Variance $ / Variance % / Variance bps | Forced-sign variance formats &mdash; pair with &sect;4.2 sign convention |
| FBM Link | Green text &mdash; cross-sheet links |
| FBM External | Red text &mdash; external file links |
| FBM Assumption | Yellow fill, blue bold &mdash; key assumption |
| FBM Total | Top single + bottom double border, bold |
| FBM Date | mm/dd/yyyy |
| FBM Year | Text format, centered |
| FBM Band | Banding grey fill (`#F7F7F7`) &mdash; alternating data rows |
| FBM KPI Big | 24pt Rockwell bold navy, centered |
| FBM KPI Label | 9pt italic charcoal, centered |

## Implementation

### Cross-platform (Python + openpyxl) — preferred for cloud / Linux / batch

Use `scripts/apply_styles.py` (bundled with this skill):

```python
import sys
sys.path.insert(0, '<path-to-skill>/scripts')
from apply_styles import inject_fbm_styles, FBM

# Load or create target workbook
from openpyxl import load_workbook
wb = load_workbook('target.xlsx')

# Inject all FBM named styles (25 in v1.2)
inject_fbm_styles(wb)

# Apply styles by name
ws = wb.active
ws['B2'].style = 'FBM Title'
ws['B2'] = 'My Report'
ws['B5'].style = 'FBM Header'
ws['C6'].style = 'FBM Input $'
ws['C6'] = 1000000

# Brand colors are available as constants
print(FBM.NAVY)   # '093254'
print(FBM.FMT_CURRENCY)  # number format string

wb.save('target.xlsx')
```

### Windows + Excel COM (PowerShell) — preferred for interactive one-offs

See `scripts/build-template.ps1` for a complete working example covering every style, named range, print setup, freeze panes, autofilter, and conditional formatting (icon set on margin %).

## Updating the standard

The skill is self-contained. To change a rule:

1. Edit `references/standards.md` (canonical text).
2. Edit `assets/template.xlsx` directly, OR edit `scripts/build-template.ps1` and re-run on Windows to regenerate.
3. Update the relevant `FBM_*` constant in `scripts/apply_styles.py`.
4. Optionally mirror the change in this SKILL.md's Quick Reference section.
5. Re-package the skill (`python -m scripts.package_skill <skill-path>`) and re-upload to Claude.ai if you maintain a cloud copy.

## Triggering reminders

- Trigger on any FBM context — workbook subject, file path containing "FBM" or "Foundation Building Materials", working folder under `OneDrive/Foundation Building Materials`, mentions of FP&A reviews / RVP / Leadership Update / REC Leases / commission reports.
- Don't require the user to explicitly say "FBM" — if they share an Excel file that's clearly in this context, apply the standard.
- When in doubt, ask once whether to apply FBM formatting.
