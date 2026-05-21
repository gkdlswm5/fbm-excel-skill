# fbm-excel

A Claude Skill that applies the **Foundation Building Materials (FBM) Excel formatting standard** — brand colors, fonts, cell styles, number formats, tab conventions, named ranges, print setup, and file naming — to any `.xlsx` workbook.

Built for FBM Operations FP&A but useful as a template for any organization wanting to standardize Excel output through Claude.

## What's included

| Path | Purpose |
|---|---|
| `SKILL.md` | Skill definition: triggering, workflow, quick reference |
| `references/standards.md` | Full canonical standard — read this first |
| `assets/template.xlsx` | Working template with 15 pre-built cell styles, named ranges, example sheets (Cover → Inputs → Calc → Output → Reference) |
| `assets/fbm-logo.jpeg` | Official FBM logo |
| `scripts/apply_styles.py` | Cross-platform openpyxl helper to inject FBM named styles into any workbook |
| `scripts/build-template.ps1` | Windows + Excel COM script that builds `template.xlsx` from scratch |
| `fbm-excel.skill` | Packaged single-file distribution for Claude.ai upload |

## Install

### Claude.ai (web/mobile)

1. Go to [claude.ai](https://claude.ai) → **Settings** → **Capabilities** → **Skills**
2. Click **Upload skill**
3. Select `fbm-excel.skill` from this repo
4. Once uploaded, the skill is available in any Claude.ai conversation across web/mobile/desktop

### Claude Code (local)

Clone or copy this repo's contents into your local skills folder:

```bash
# macOS/Linux
git clone https://github.com/gkdlswm5/fbm-excel-skill.git ~/.claude/skills/fbm-excel
```

```powershell
# Windows PowerShell
git clone https://github.com/gkdlswm5/fbm-excel-skill.git $env:USERPROFILE\.claude\skills\fbm-excel
```

New Claude Code sessions will pick up the skill automatically.

## How to use

Once installed, just ask Claude to apply the standard:

- "Apply the FBM standard to this file"
- "Reformat this spreadsheet to FBM"
- "Make this Excel file look on-brand"
- "Standardize this workbook"

Or invoke explicitly via `/fbm-excel` in Claude Code.

## Brand standard summary

| Element | Convention |
|---|---|
| Primary colors | Navy `#093254`, burgundy `#6A1831`, forest `#005E34` |
| Fonts | Rockwell (headers), Calibri 9 (body) |
| Tab colors | Blue = inputs · Green = outputs · Gray = reference · Navy = cover |
| Cell text colors | Blue input · Black formula · Green cross-sheet link · Red external link |
| Negatives | Parentheses, not minus signs |
| File naming | `FBM - [Subject] - [YYYY.MM.DD].xlsx` |

Full rules in [`references/standards.md`](references/standards.md).

## Updating the standard

The single source of truth is `references/standards.md`. When you change a rule:

1. Edit `references/standards.md`
2. Update `assets/template.xlsx` (re-run `scripts/build-template.ps1` on Windows, or edit the .xlsx directly)
3. Update the relevant constant in `scripts/apply_styles.py` if it changes a color or format
4. Re-package for cloud: `python -m scripts.package_skill <path-to-skill-folder>` (requires `skill-creator`)
5. Re-upload `fbm-excel.skill` to Claude.ai (overwrite the existing one)

## License

Internal FBM use. Brand assets (colors, fonts, logo) are property of Foundation Building Materials. Code is MIT.
