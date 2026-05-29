"""Apply FBM Excel formatting standard to a workbook via openpyxl.

Cross-platform helper for the fbm-excel skill. Use this in cloud environments
(Claude.ai) or anywhere Excel COM isn't available.

Usage:
    from apply_styles import inject_fbm_styles, FBM
    from openpyxl import load_workbook

    wb = load_workbook('target.xlsx')
    inject_fbm_styles(wb)

    ws = wb.active
    ws['B2'].style = 'FBM Title'
    ws['B5'].style = 'FBM Header'
    ws['C6'].style = 'FBM Input $'
    wb.save('target.xlsx')
"""
from openpyxl.styles import (
    NamedStyle, Font, PatternFill, Alignment, Border, Side, Color
)


class FBM:
    """FBM brand constants — colors, fonts, number formats."""

    # Brand colors (ARGB hex strings for openpyxl, no '#' prefix)
    NAVY       = '093254'
    BURGUNDY   = '6A1831'
    SLATE_BLUE = '4E7087'
    CHARCOAL   = '636466'
    FOREST     = '005E34'
    SAGE       = '708573'
    LIGHT_GREY = 'E7E6E6'
    WHITE      = 'FFFFFF'
    BLACK      = '000000'

    # Cell text colors (convention)
    INPUT_BLUE   = '0000FF'  # hardcoded inputs
    FORMULA_BLK  = '000000'  # formulas
    LINK_GREEN   = '008000'  # cross-sheet links
    EXTERNAL_RED = 'FF0000'  # external file links
    ASSUMPTION_FILL = 'FFFF00'  # yellow fill for key assumptions

    # Tab colors
    TAB_BLUE  = '1F4E79'  # Inputs
    TAB_GREEN = '005E34'  # Outputs
    TAB_GRAY  = 'A5A5A5'  # Reference
    TAB_NAVY  = '093254'  # Cover

    # Conditional formatting palette
    GOOD_FONT, GOOD_FILL       = '006100', 'C6EFCE'
    NEUTRAL_FONT, NEUTRAL_FILL = '9C5700', 'FFEB9C'
    BAD_FONT, BAD_FILL         = '9C0006', 'FFC7CE'

    # Number format strings (single standard each)
    FMT_CURRENCY  = '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'
    FMT_CURRENCY2 = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
    FMT_NUMBER    = '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'
    FMT_PERCENT   = '0.0%;(0.0%);"-"'
    FMT_MULTIPLE  = '0.0"x"'
    FMT_DATE      = 'mm/dd/yyyy'
    FMT_YEAR      = '@'

    # Fonts
    HEADER_FONT_NAME = 'Rockwell'   # fallback to Calibri Bold if unavailable
    BODY_FONT_NAME   = 'Calibri'
    BODY_FONT_SIZE   = 9
    HEADER_FONT_SIZE = 11
    TITLE_FONT_SIZE  = 14


def _color(hex_no_hash: str) -> Color:
    """openpyxl Color from a 6-char ARGB hex string."""
    return Color(rgb='FF' + hex_no_hash)


def _make_style(name: str, **kwargs) -> NamedStyle:
    """Build a NamedStyle with the given attributes."""
    s = NamedStyle(name=name)
    if 'font' in kwargs:
        s.font = kwargs['font']
    if 'fill' in kwargs:
        s.fill = kwargs['fill']
    if 'alignment' in kwargs:
        s.alignment = kwargs['alignment']
    if 'border' in kwargs:
        s.border = kwargs['border']
    if 'number_format' in kwargs:
        s.number_format = kwargs['number_format']
    return s


def inject_fbm_styles(wb) -> None:
    """Add all 15 FBM named styles to a workbook.

    Safe to call multiple times — skips styles that already exist.
    """
    existing = set(wb.named_styles)

    def add(style: NamedStyle) -> None:
        if style.name not in existing:
            wb.add_named_style(style)

    body = lambda **kw: Font(name=FBM.BODY_FONT_NAME, size=FBM.BODY_FONT_SIZE, **kw)

    # Title (Rockwell 14 navy)
    add(_make_style('FBM Title',
        font=Font(name=FBM.HEADER_FONT_NAME, size=FBM.TITLE_FONT_SIZE,
                  bold=True, color=_color(FBM.NAVY))))

    # Header (navy fill, white bold, centered)
    add(_make_style('FBM Header',
        font=Font(name=FBM.BODY_FONT_NAME, size=FBM.HEADER_FONT_SIZE,
                  bold=True, color=_color(FBM.WHITE)),
        fill=PatternFill('solid', fgColor=FBM.NAVY),
        alignment=Alignment(horizontal='center', vertical='center'),
        border=Border(*(Side(style='thin') for _ in range(4)))))

    # Subheader (light grey fill, navy bold)
    add(_make_style('FBM Subheader',
        font=Font(name=FBM.BODY_FONT_NAME, size=10, bold=True,
                  color=_color(FBM.NAVY)),
        fill=PatternFill('solid', fgColor=FBM.LIGHT_GREY),
        alignment=Alignment(horizontal='left'),
        border=Border(*(Side(style='thin') for _ in range(4)))))

    # Input styles (blue text)
    add(_make_style('FBM Input',
        font=body(color=_color(FBM.INPUT_BLUE)),
        number_format=FBM.FMT_NUMBER))
    add(_make_style('FBM Input $',
        font=body(color=_color(FBM.INPUT_BLUE)),
        number_format=FBM.FMT_CURRENCY))
    add(_make_style('FBM Input %',
        font=body(color=_color(FBM.INPUT_BLUE)),
        number_format=FBM.FMT_PERCENT))

    # Formula styles (black text)
    add(_make_style('FBM Formula',
        font=body(color=_color(FBM.FORMULA_BLK)),
        number_format=FBM.FMT_NUMBER))
    add(_make_style('FBM Formula $',
        font=body(color=_color(FBM.FORMULA_BLK)),
        number_format=FBM.FMT_CURRENCY))
    add(_make_style('FBM Formula %',
        font=body(color=_color(FBM.FORMULA_BLK)),
        number_format=FBM.FMT_PERCENT))

    # Cross-sheet link (green text)
    add(_make_style('FBM Link',
        font=body(color=_color(FBM.LINK_GREEN)),
        number_format=FBM.FMT_NUMBER))

    # External link (red text)
    add(_make_style('FBM External',
        font=body(color=_color(FBM.EXTERNAL_RED)),
        number_format=FBM.FMT_NUMBER))

    # Key assumption (yellow fill, blue bold)
    add(_make_style('FBM Assumption',
        font=body(color=_color(FBM.INPUT_BLUE), bold=True),
        fill=PatternFill('solid', fgColor=FBM.ASSUMPTION_FILL),
        number_format='General'))

    # Total row (top single + bottom double border)
    add(_make_style('FBM Total',
        font=body(color=_color(FBM.FORMULA_BLK), bold=True),
        border=Border(top=Side(style='thin'), bottom=Side(style='double')),
        number_format=FBM.FMT_NUMBER))

    # Date
    add(_make_style('FBM Date',
        font=body(color=_color(FBM.FORMULA_BLK)),
        number_format=FBM.FMT_DATE))

    # Year (text, centered)
    add(_make_style('FBM Year',
        font=body(color=_color(FBM.FORMULA_BLK)),
        alignment=Alignment(horizontal='center'),
        number_format=FBM.FMT_YEAR))


def apply_standard_layout(
    ws,
    header_row: int = 5,
    title_text: str = '',
    subtitle_text: str = '',
    label_col: str = 'B',
) -> None:
    """Apply standard FBM layout to a worksheet:
      - Col A = 5 (gutter)
      - Col B = 30 (labels)
      - Other cols = 15
      - Title in B2 with FBM Title style
      - Optional subtitle in B3 (italic charcoal, wrap OFF, overflow right)
      - Freeze panes computed from header_row and label_col

    Freeze pane is placed one row below header_row and one column right of
    label_col. Defaults (header_row=5, label_col='B') resolve to C6. If the
    header sits in row 7 and labels span B:C, pass label_col='C' to freeze
    at D8.
    """
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 30
    for col_letter in 'CDEFGHIJ':
        ws.column_dimensions[col_letter].width = 15

    if title_text:
        ws['B2'] = title_text
        ws['B2'].style = 'FBM Title'

    if subtitle_text:
        set_subtitle(ws, subtitle_text)

    ws.freeze_panes = _freeze_anchor(header_row, label_col)


def set_subtitle(ws, text: str, row: int = 3, col: str = 'B') -> None:
    """Write an italic charcoal subtitle that does NOT wrap.

    Wrapping a long sentence inside the 30-width label column inflates the
    row height. This helper forces wrap_text=False and left alignment so
    long text overflows into the empty cells to the right.
    """
    cell = ws[f'{col}{row}']
    cell.value = text
    cell.font = Font(
        name=FBM.BODY_FONT_NAME,
        size=FBM.BODY_FONT_SIZE,
        italic=True,
        color=_color(FBM.CHARCOAL),
    )
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)


def disable_pivot_autofit(wb) -> int:
    """For every pivot table in the workbook: disable autofit-on-refresh and
    enable preserve-formatting-on-refresh.

    Maps to the PivotTable Options checkboxes:
      - "Autofit column widths on update" -> OFF (useAutoFormatting=False)
      - "Preserve cell formatting on update" -> ON (preserveFormatting=True)

    Returns the count of pivot tables touched.
    """
    count = 0
    for ws in wb.worksheets:
        for pivot in getattr(ws, '_pivots', []):
            pivot.useAutoFormatting = False
            pivot.preserveFormatting = True
            count += 1
    return count


def _freeze_anchor(header_row: int, label_col: str) -> str:
    """Return the freeze-pane anchor: one row below header, one col right of label_col."""
    n = 0
    for ch in label_col.upper():
        n = n * 26 + (ord(ch) - ord('A') + 1)
    n += 1
    next_col = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        next_col = chr(r + ord('A')) + next_col
    return f'{next_col}{header_row + 1}'


def set_tab_color(ws, role: str) -> None:
    """Set worksheet tab color per FBM convention.

    role: 'input', 'output', 'reference', or 'cover'
    """
    colors = {
        'input':     FBM.TAB_BLUE,
        'output':    FBM.TAB_GREEN,
        'reference': FBM.TAB_GRAY,
        'cover':     FBM.TAB_NAVY,
    }
    if role not in colors:
        raise ValueError(f"role must be one of {list(colors)}")
    ws.sheet_properties.tabColor = colors[role]


if __name__ == '__main__':
    # Self-test: build a tiny demo workbook
    from openpyxl import Workbook
    wb = Workbook()
    inject_fbm_styles(wb)

    ws = wb.active
    ws.title = 'Demo'
    apply_standard_layout(
        ws,
        title_text='FBM Style Demo',
        subtitle_text='One row per (RawData x matching rule). Long subtitles overflow right, they do not wrap.',
    )
    set_tab_color(ws, 'output')

    # Header row
    for col, label in enumerate(['Item', 'FY2026', 'FY2027', 'FY2028'], start=2):
        cell = ws.cell(row=5, column=col, value=label)
        cell.style = 'FBM Header'

    # Sample input row
    ws['B6'] = 'Revenue'
    for col, val in enumerate([10_000_000, 10_500_000, 11_025_000], start=3):
        cell = ws.cell(row=6, column=col, value=val)
        cell.style = 'FBM Input $'

    # Sample formula row
    ws['B7'] = 'Growth %'
    ws['D7'] = '=D6/C6-1'
    ws['E7'] = '=E6/D6-1'
    for col_letter in ['D', 'E']:
        ws[f'{col_letter}7'].style = 'FBM Formula %'

    wb.save('fbm_demo.xlsx')
    print('Saved fbm_demo.xlsx')
