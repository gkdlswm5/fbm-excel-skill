# FBM Excel Standard — Polish Roadmap

**Owner:** Andrew Kim &middot; Operations FP&A &middot; Last updated: 2026-05-29

This document lists recommended improvements to bring FBM workbooks to an S&P 500 / investor-relations level of polish. The current `standards.md` is v1.0; these items are queued for a future v1.1 / v1.2 / v1.3.

Items are grouped into three implementation waves (PR-A &rarr; PR-B &rarr; PR-C) so they can be rolled out without rewriting the standard in one go. Pick a wave, say the word, and it gets executed against `standards.md`, `apply_styles.py`, `build-template.ps1`, and `assets/template.xlsx`.

> **Already incorporated (not in any wave):**
> - "Positive good, negative bad" sign convention &mdash; landed in `standards.md` &sect;4.2 on 2026-05-29. Recommendation #1 below is therefore partially done; the sign-convention half is shipped, the cost/COGS display-side wording stays as future work.

---

## PR-A &mdash; Foundation polish

Pure standards / template / helper changes. No new visual elements. Lowest risk, highest baseline-quality gain.

### A1. Sign convention &mdash; positive good, negative bad
**Status: SHIPPED in &sect;4.2.** Remaining sub-work: document COGS / expense display convention (positive values for costs unless rebate / contra-revenue). Add to &sect;3.

### A2. Variance / delta number formats
Add three new number-format standards to &sect;3.1:
- Variance ($): `+#,##0;(#,##0);"-"`
- Variance (%): `+0.0%;(0.0%);"-"`
- bps delta: `+0" bps";(0)" bps";"-"`

Add matching named styles to `apply_styles.py` and `build-template.ps1`: `FBM Variance`, `FBM Variance $`, `FBM Variance %`, `FBM Variance bps`.

### A3. Units row under headers
Reserve row 4 (between subtitle and header) for unit labels &mdash; `$ thousands`, `$ millions`, `units`, `%`, `bps`. Italic charcoal, centered. Eliminates the most common reviewer question. Update &sect;2.2 layout convention and the freeze-pane default (`C6` becomes `C7` if row 4 is in use).

### A4. Scale discipline
Pick one scale per workbook (`$ thousands` OR `$ millions`) and declare it on the Cover sheet. Never mix within a sheet. Add a named range `_Scale` and a Cover-sheet field "Reporting scale". Reference in any printed header (e.g. `&"Calibri,Bold"&12Inputs ($ thousands)`).

### A5. Color discipline tightened
Lock the role of each brand color so analysts can't free-style:
- Navy &mdash; primary headers, sheet titles, cover tab.
- Slate blue &mdash; secondary / grouping headers within a section.
- Burgundy &mdash; callouts only (warnings, footnotes, "audit issue").
- Sage &mdash; soft category accents (chart fills, sub-totals).
- Light grey &mdash; banding fill on alternating data rows.
- Forest green &mdash; output tab, "shipped" / "approved" status only.
- Charcoal &mdash; body text, muted labels, gridline-like borders.

Forbid ad-hoc colors. Add the matrix to &sect;1.1.

### A6. Banding on data tables
Subtle `#F7F7F7` fill on alternating data rows (rows 6, 8, 10, &hellip;). Apply via Format-as-Table or a conditional format using `=MOD(ROW(),2)=0`. Add a `FBM Band` style and update the template to demonstrate.

### A7. Row-height and column-width hygiene
Lock standard row heights so every sheet looks identical:
- Row 2 (title): 24pt
- Row 3 (subtitle): 16pt
- Row 4 (units, when present): 16pt
- Row 5 (header): 30pt
- Row 6+ (data): 16pt
- Row above totals: 8pt (blank breathing-room row)

Column widths already covered in &sect;2.2 (5 / 30 / 15) &mdash; reinforce that multiples of 5 are required.

### A8. Indent sub-line-items via alignment, not spaces
Sub-items use `Alignment(indent=N)` not leading spaces. Document and add a helper `set_indent(cell, level)`.

### A9. Right-align numeric headers, center text headers, left-align labels
Year columns like `FY2026` are numeric &mdash; right-align. Text headers (e.g. `Region`) &mdash; center. Label column B &mdash; always left. Document explicitly.

### A10. Decimal-place consistency per column
All `$` values in the same column use the same precision &mdash; never mix `$1,234` and `$1,234.56`. Add as a hard rule in &sect;3.

### A11. Total rows have a blank row above them
Visual breathing room. The single-thin top border on `FBM Total` does double duty when there's an 8pt blank row above. Document and update template examples.

### A12. Expand named ranges
The current template defines four (`Revenue_Input`, `Growth_Rate`, `Tax_Rate`, `Discount_Rate`). Expand to: `AsOfDate`, `Scale`, `ReportingCurrency`, `FY_Start`, `FY_End`, `Discount_Rate`, `Tax_Rate`, `WACC`, `Inflation_Rate`. Update `build-template.ps1`.

### A13. Workbook properties
Set File &rarr; Info &rarr; Properties: `Title`, `Subject`, `Author`, `Company = Foundation Building Materials`, `Category`. Surfaces in DMS and SharePoint. Add to `build-template.ps1` and as a checklist item.

---

## PR-B &mdash; Output polish

Visible upgrades to deliverable-facing sheets (Cover, Output). Higher visual impact than PR-A; lower risk than PR-C.

### B1. Hide gridlines on output-facing sheets
Cover and Output: gridlines off. Inputs and Calc: gridlines on (helps the analyst). Set via `ws.sheet_view.showGridLines = False` / `$ws.DisplayGridlines = $false`. Document in &sect;2 and update the template.

### B2. Cover sheet upgrade
Add to the Cover layout:
- As-of date field (named: `AsOfDate`)
- Reporting currency + scale fields (named: `ReportingCurrency`, `Scale`)
- Contact email field
- "Data sources" block (3-5 rows: source system, last refreshed, owner)
- Version-history mini-table (5 most recent entries: version, date, author, note)
- Confidentiality footer: `FBM Confidential &mdash; Do Not Distribute`

### B3. Footer standard tightened
Add workbook version and as-of date to the print footer. Updated &sect;4.4 (Print setup) row:
- Left: `&F` (file path)
- Center: `&D` (system date) + " &middot; v" + Workbook.Version + " &middot; As of " + AsOfDate
- Right: `&P of &N`

### B4. KPI cards on Output
Boxed metric tiles with big number + tiny variance below. Two new named styles:
- `FBM KPI Big` &mdash; 24pt Rockwell navy, centered.
- `FBM KPI Label` &mdash; 9pt charcoal, centered, italic.

Add an example KPI card row to the Output sheet of the template.

### B5. Sparklines / trend mini-charts on Output
Any time-series row in Output gets a column with a line sparkline. One-line addition (`ws.SparklineGroups.Add(...)`), large professional impact.

### B6. Hyperlinked TOC on the Cover
Add a small TOC block to the Cover sheet:
- `=HYPERLINK("#Inputs!A1","Inputs")`
- `=HYPERLINK("#Calc!A1","Calc")`
- `=HYPERLINK("#Output!A1","Output")`
- `=HYPERLINK("#Reference!A1","Reference")`

Style: `FBM Link` (green), one per row in the cover metadata block.

---

## PR-C &mdash; Rigor

Behavioral / process / control changes. Highest payoff for review-readiness; affects how analysts work, not just how the file looks.

### C1. Conditional-formatting rules of engagement
Document which columns get CF and which don't:
- Variance %, growth %, margin % &mdash; always get the good/neutral/bad palette (CF since sign convention guarantees green=good).
- Raw $ &mdash; never get CF (only variance from $ does).
- Headcount &mdash; never CF (context-dependent).
Add explicit examples to &sect;3.5.

### C2. Number-format edge cases
Add to &sect;3.1:
- Ratio: `0.00`
- Share count: `#,##0`
- Share price: `#,##0.00`
- Basis points: `0" bps"` (and variance bps: see A2)
- Non-applicable cells: literal `"n/a"` in the cell, not blank, not zero, not `#N/A`.

### C3. Print fit floor
Current rule is "fit to 1 page wide, unlimited tall". Add: scale floor of 60%. Anything that would compress below 60% gets reorganized (split sheet, narrower columns), never shrunk to fit. Document in &sect;4.4.

### C4. Data validation on every editable cell
Every blue (input) cell must have data validation: type, range, or list. Document the convention and the rule that PR-stage review will reject Inputs sheets without it.

### C5. Sheet protection on Calc and Output
Lock formula cells; allow editing only on Input cells. Default protection on for Calc and Output; password optional. Document the steps in &sect;4.6.

### C6. Negatives stay charcoal (not red)
Reinforce: red is reserved for external links only. Variance columns use the good/neutral/bad CF palette &mdash; the unfavorable cells get a red fill (`#FFC7CE`) and dark-red font (`#9C0006`) from the palette, not generic black-on-red. Document explicitly so analysts don't switch the negative-format string to a red one.

### C7. Spell-check before save
Add to the &sect;5 quick checklist. `F7` in Excel.

### C8. Print preview every sheet before sharing
Add to the &sect;5 quick checklist. Catches page-break errors that other checks miss.

---

## Implementation notes

When you're ready to execute a wave:

1. **Pick the wave** (`PR-A`, `PR-B`, or `PR-C`) or individual items.
2. **Branch** off the current `main` (current dev branch: `claude/awesome-ramanujan-1QEi5`).
3. **Update in order:**
   - `references/standards.md` (canonical doc)
   - `scripts/apply_styles.py` (Python helpers + named styles)
   - `scripts/build-template.ps1` (Windows + Excel COM template builder)
   - `assets/template.xlsx` (regenerate by running `build-template.ps1` on Windows)
   - `SKILL.md` Quick Reference section
4. **Add checklist items** to &sect;5 of `standards.md` for any new manual-verification rules.
5. **Bump the version** at the top of `standards.md` (v1.0 &rarr; v1.1 after PR-A, v1.2 after PR-B, v1.3 after PR-C).
