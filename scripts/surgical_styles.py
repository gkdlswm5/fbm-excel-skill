"""Surgical FBM style injection for openpyxl-unsafe workbooks.

Use this when ``apply_styles.inject_fbm_styles`` is unsafe — workbooks containing
PivotTables, charts, external connections, or custom XML parts that openpyxl
mangles on save. This helper performs **byte-level insertion** into the
existing ``xl/styles.xml``; it never reserializes the ``styleSheet`` root
element, so the XML declaration quote style, line endings, and root-attribute
order remain byte-identical to what Excel originally wrote.

Why byte-level?
================

Excel's XML parser is **stricter than the OOXML spec**. Three lxml defaults
that the spec permits but Excel rejects when applied to ``styles.xml`` or
``workbook.xml``:

1. XML declaration quote style — Excel requires double quotes
   ``<?xml version="1.0"`` (lxml writes single).
2. Line ending after ``?>`` — Excel requires ``\r\n`` (lxml writes ``\n``).
3. Attribute order on the root element — Excel requires the order it
   originally wrote (lxml reorders).

Violating any one of these makes Excel discard the whole ``styles.xml`` part
on open. Once that happens, every cell's ``s="N"`` reference becomes invalid
and the cascade hits every formatted sheet in the workbook (sometimes
PivotTables too).

The functions in this module avoid the trap by inserting new entries before
each section's closing tag via regex, and updating the section's ``count``
attribute in place. The root element and the XML declaration are never
touched.

Public API
==========

``extend_styles_bytewise(styles_bytes) -> (new_bytes, xf_index_map)``
    Append all FBM v1.3 styles to styles.xml. Returns the modified bytes and
    a dict mapping style name (e.g. ``"FBM Title"``) to its new ``cellXfs``
    index. Idempotent if called once — calling twice will add a second copy
    of each style, so guard with your own check.

``apply_cell_styles(sheet_bytes, coord_to_xfid) -> new_bytes``
    Update the ``s="N"`` attribute on specific cells in a worksheet XML.
    Operates via lxml on the sheet body (Excel is lenient about sheet
    reserialization, just not about workbook-level parts).

``set_tab_color(sheet_bytes, hex_rgb) -> new_bytes``
    Set or replace the tab color on a sheet via lxml. Safe on sheet bodies.

Typical workflow
================

::

    import zipfile, shutil
    from surgical_styles import extend_styles_bytewise, apply_cell_styles, set_tab_color, FBM_TAB

    shutil.copy2(src, work)
    with zipfile.ZipFile(work, 'r') as zin:
        parts = {info.filename: zin.read(info.filename) for info in zin.infolist()}

    # 1. Inject FBM styles into styles.xml (byte-level — preserves Excel invariants)
    parts['xl/styles.xml'], xf = extend_styles_bytewise(parts['xl/styles.xml'])

    # 2. Resolve sheet paths from workbook.xml (don't hardcode — Excel renumbers)
    #    ... use your own helper, or see resolve_sheet_paths() below.

    # 3. Apply styles to specific cells per your layout map
    cfg_styles = {
        'A1': xf['FBM Title'],
        'A3': xf['FBM Subheader'],
        'A8': xf['FBM Header'], 'B8': xf['FBM Header'], ...
    }
    parts['xl/worksheets/sheet11.xml'] = apply_cell_styles(
        parts['xl/worksheets/sheet11.xml'], cfg_styles)

    # 4. Tab colors
    parts['xl/worksheets/sheet11.xml'] = set_tab_color(
        parts['xl/worksheets/sheet11.xml'], FBM_TAB['input'])

    # 5. Drop calcChain so Excel recomputes (optional but recommended)
    parts.pop('xl/calcChain.xml', None)

    # 6. Repack
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fn, data in parts.items():
            zout.writestr(fn, data)
"""
from __future__ import annotations

import re
from typing import Dict, Mapping, Tuple


# ─── FBM v1.3 brand constants (mirror of apply_styles.FBM) ──────────────────
FBM_NAVY      = "093254"
FBM_BURGUNDY  = "6A1831"
FBM_CHARCOAL  = "636466"
FBM_FOREST    = "005E34"
FBM_LT_GREY   = "E7E6E6"
FBM_BAND_GREY = "F7F7F7"
FBM_WHITE     = "FFFFFF"

FBM_INPUT_BLUE       = "0000FF"
FBM_FORMULA_BLK      = "000000"
FBM_LINK_GREEN       = "008000"
FBM_EXTERNAL_RED     = "FF0000"
FBM_ASSUMPTION_FILL  = "FFFF00"

# Tab colors per role (use with set_tab_color)
FBM_TAB = {
    "input":     "1F4E79",
    "output":    "005E34",
    "reference": "A5A5A5",
    "cover":     "093254",
}

# Number formats — ``&quot;`` is the XML-escaped form of ``"``
_FMT_CURRENCY     = '_($* #,##0_);_($* (#,##0);_($* &quot;-&quot;_);_(@_)'
_FMT_CURRENCY2    = '_($* #,##0.00_);_($* (#,##0.00);_($* &quot;-&quot;??_);_(@_)'
_FMT_NUMBER       = '_(* #,##0_);_(* (#,##0);_(* &quot;-&quot;_);_(@_)'
_FMT_PERCENT      = '0.0%;(0.0%);&quot;-&quot;'
_FMT_DATE         = 'mm/dd/yyyy'
_FMT_YEAR         = '@'
_FMT_VARIANCE     = '+#,##0;(#,##0);&quot;-&quot;'
_FMT_VARIANCE_DOL = '+$#,##0;($#,##0);&quot;-&quot;'
_FMT_VARIANCE_PCT = '+0.0%;(0.0%);&quot;-&quot;'
_FMT_VARIANCE_BPS = '+0&quot; bps&quot;;(0)&quot; bps&quot;;&quot;-&quot;'


# ─── Internal: byte-level section helpers ───────────────────────────────────

def _section_close(text: str, tag: str) -> int:
    """Return the byte offset just before ``</tag>``.

    Raises ``ValueError`` if the section isn't found. ``styles.xml`` always
    has fonts/fills/borders/cellXfs from Excel's default save, so this is
    a hard error.
    """
    open_match = re.search(rf'<{tag}\b', text)
    if open_match is None:
        raise ValueError(f"No <{tag}> section in styles.xml")
    close = text.find(f"</{tag}>", open_match.end())
    if close < 0:
        raise ValueError(f"Unclosed <{tag}> in styles.xml")
    return close


def _section_count(text: str, tag: str) -> int:
    """Return the current ``count`` attribute on a section, or 0 if absent."""
    m = re.search(rf'<{tag}\b[^>]*?count="(\d+)"', text)
    return int(m.group(1)) if m else 0


def _bump_count(text: str, tag: str, delta: int) -> str:
    """Increase the ``count`` attribute on a section by ``delta``.

    If the section has no ``count`` attribute (rare), add one.
    """
    m = re.search(rf'(<{tag}\b)([^>]*?)>', text)
    if m is None:
        return text
    head, attrs = m.group(1), m.group(2)
    count_match = re.search(r'count="(\d+)"', attrs)
    if count_match is None:
        new_attrs = f'{attrs} count="{delta}"'
    else:
        new_count = int(count_match.group(1)) + delta
        new_attrs = (
            attrs[:count_match.start()]
            + f'count="{new_count}"'
            + attrs[count_match.end():]
        )
    return text[:m.start()] + head + new_attrs + ">" + text[m.end():]


def _next_numfmt_id(text: str) -> int:
    """Find the next free ``numFmtId`` value for custom formats.

    Excel reserves IDs 0–163 for built-ins, and most existing files use
    164–199 for one-off custom formats. We start at 200 to avoid collision.
    """
    ids = {int(x) for x in re.findall(r'numFmtId="(\d+)"', text)}
    return max(200, max(ids, default=199) + 1)


# ─── Public API ─────────────────────────────────────────────────────────────

def extend_styles_bytewise(
    styles_bytes: bytes,
) -> Tuple[bytes, Dict[str, int]]:
    """Append FBM v1.3 styles to a workbook's ``xl/styles.xml``.

    Inserts new ``numFmts``, ``fonts``, ``fills``, ``borders``, and
    ``cellXfs`` entries before each section's closing tag. Updates the
    ``count`` attribute on each section. The ``styleSheet`` root element
    and the XML declaration are untouched — byte-for-byte identical to
    input.

    Returns
    -------
    new_bytes : bytes
        Modified styles.xml content.
    xf_index_map : dict
        Maps style name (e.g. ``"FBM Title"``) to its new ``cellXfs`` index.
        Use these indexes when setting cell ``s="N"`` attributes.
    """
    text = styles_bytes.decode("utf-8")

    # ── numFmts ─────────────────────────────────────────────────────────────
    next_nf = _next_numfmt_id(text)
    nf_specs = [
        ("currency",     _FMT_CURRENCY),
        ("currency2",    _FMT_CURRENCY2),
        ("number",       _FMT_NUMBER),
        ("percent",      _FMT_PERCENT),
        ("date",         _FMT_DATE),
        ("year",         _FMT_YEAR),
        ("variance",     _FMT_VARIANCE),
        ("variance_dol", _FMT_VARIANCE_DOL),
        ("variance_pct", _FMT_VARIANCE_PCT),
        ("variance_bps", _FMT_VARIANCE_BPS),
    ]
    nf_idx: Dict[str, int] = {}
    nf_xml = []
    for key, fmt in nf_specs:
        nf_xml.append(f'<numFmt numFmtId="{next_nf}" formatCode="{fmt}"/>')
        nf_idx[key] = next_nf
        next_nf += 1
    if "<numFmts" in text:
        close = _section_close(text, "numFmts")
        text = text[:close] + "".join(nf_xml) + text[close:]
        text = _bump_count(text, "numFmts", len(nf_xml))
    else:
        # No numFmts section exists; insert one right after the styleSheet open
        root_close = text.find(">", text.find("<styleSheet")) + 1
        block = f'<numFmts count="{len(nf_xml)}">' + "".join(nf_xml) + "</numFmts>"
        text = text[:root_close] + block + text[root_close:]

    # ── fonts ───────────────────────────────────────────────────────────────
    fonts_count_before = _section_count(text, "fonts")
    font_specs = []   # list of (key, xml)
    def _font(key, name="Calibri", sz=9, bold=False, italic=False, color=FBM_FORMULA_BLK):
        parts = [f'<sz val="{sz}"/>', f'<color rgb="FF{color}"/>',
                 f'<name val="{name}"/>', '<family val="2"/>']
        if bold:   parts.append('<b/>')
        if italic: parts.append('<i/>')
        font_specs.append((key, '<font>' + ''.join(parts) + '</font>'))

    _font("title",        name="Rockwell", sz=14, bold=True, color=FBM_NAVY)
    _font("subtitle",     italic=True, color=FBM_CHARCOAL)
    _font("units",        italic=True, color=FBM_CHARCOAL)
    _font("header",       sz=10, bold=True, color=FBM_WHITE)
    _font("subheader",    sz=10, bold=True, color=FBM_NAVY)
    _font("body_black",   color=FBM_FORMULA_BLK)
    _font("body_blue",    color=FBM_INPUT_BLUE)
    _font("body_green",   color=FBM_LINK_GREEN)
    _font("body_red",     color=FBM_EXTERNAL_RED)
    _font("body_blue_b",  bold=True, color=FBM_INPUT_BLUE)
    _font("body_black_b", bold=True, color=FBM_FORMULA_BLK)
    _font("kpi_big",      name="Rockwell", sz=24, bold=True, color=FBM_NAVY)

    font_idx: Dict[str, int] = {key: fonts_count_before + i for i, (key, _) in enumerate(font_specs)}
    close = _section_close(text, "fonts")
    text = text[:close] + "".join(xml for _, xml in font_specs) + text[close:]
    text = _bump_count(text, "fonts", len(font_specs))

    # ── fills ───────────────────────────────────────────────────────────────
    fills_count_before = _section_count(text, "fills")
    fill_specs = []
    def _fill(key, rgb):
        fill_specs.append((key,
            f'<fill><patternFill patternType="solid">'
            f'<fgColor rgb="FF{rgb}"/><bgColor indexed="64"/>'
            f'</patternFill></fill>'))
    _fill("navy", FBM_NAVY)
    _fill("lt_grey", FBM_LT_GREY)
    _fill("band_grey", FBM_BAND_GREY)
    _fill("yellow", FBM_ASSUMPTION_FILL)

    fill_idx: Dict[str, int] = {key: fills_count_before + i for i, (key, _) in enumerate(fill_specs)}
    close = _section_close(text, "fills")
    text = text[:close] + "".join(xml for _, xml in fill_specs) + text[close:]
    text = _bump_count(text, "fills", len(fill_specs))

    # ── borders ─────────────────────────────────────────────────────────────
    borders_count_before = _section_count(text, "borders")
    border_specs = [
        ("thin",
         '<border>'
         '<left style="thin"><color rgb="FF000000"/></left>'
         '<right style="thin"><color rgb="FF000000"/></right>'
         '<top style="thin"><color rgb="FF000000"/></top>'
         '<bottom style="thin"><color rgb="FF000000"/></bottom>'
         '<diagonal/></border>'),
        ("total",
         '<border>'
         '<left/><right/>'
         '<top style="thin"><color rgb="FF000000"/></top>'
         '<bottom style="double"><color rgb="FF000000"/></bottom>'
         '<diagonal/></border>'),
    ]
    border_idx: Dict[str, int] = {key: borders_count_before + i for i, (key, _) in enumerate(border_specs)}
    close = _section_close(text, "borders")
    text = text[:close] + "".join(xml for _, xml in border_specs) + text[close:]
    text = _bump_count(text, "borders", len(border_specs))

    # ── cellXfs ─────────────────────────────────────────────────────────────
    cellxfs_count_before = _section_count(text, "cellXfs")
    xf_specs = []
    def _xf(name, font_key, fill_key=None, border_key=None, numfmt_key=None,
            horiz=None, vert=None, wrap=False):
        font = font_idx[font_key]
        fill = fill_idx[fill_key] if fill_key else 0
        border = border_idx[border_key] if border_key else 0
        numfmt = nf_idx[numfmt_key] if numfmt_key else 0
        attrs = [
            f'numFmtId="{numfmt}"', f'fontId="{font}"',
            f'fillId="{fill}"', f'borderId="{border}"',
            'xfId="0"', 'applyFont="1"',
        ]
        if fill:   attrs.append('applyFill="1"')
        if border: attrs.append('applyBorder="1"')
        if numfmt: attrs.append('applyNumberFormat="1"')
        has_align = horiz or vert or wrap
        if has_align: attrs.append('applyAlignment="1"')
        opening = '<xf ' + ' '.join(attrs)
        if has_align:
            a = []
            if horiz: a.append(f'horizontal="{horiz}"')
            if vert:  a.append(f'vertical="{vert}"')
            if wrap:  a.append('wrapText="1"')
            xml = opening + '><alignment ' + ' '.join(a) + '/></xf>'
        else:
            xml = opening + '/>'
        xf_specs.append((name, xml))

    _xf("FBM Title",        "title")
    _xf("FBM Subtitle",     "subtitle", horiz="left", vert="center")
    _xf("FBM Units",        "units", horiz="center", vert="center")
    _xf("FBM Header",       "header", fill_key="navy", border_key="thin",
        horiz="center", vert="center", wrap=True)
    _xf("FBM Subheader",    "subheader", fill_key="lt_grey", border_key="thin",
        horiz="left", vert="center")
    _xf("FBM Input",        "body_blue", numfmt_key="number")
    _xf("FBM Input $",      "body_blue", numfmt_key="currency")
    _xf("FBM Input %",      "body_blue", numfmt_key="percent")
    _xf("FBM Formula",      "body_black", numfmt_key="number")
    _xf("FBM Formula $",    "body_black", numfmt_key="currency")
    _xf("FBM Formula %",    "body_black", numfmt_key="percent")
    _xf("FBM Variance",     "body_black", numfmt_key="variance")
    _xf("FBM Variance $",   "body_black", numfmt_key="variance_dol")
    _xf("FBM Variance %",   "body_black", numfmt_key="variance_pct")
    _xf("FBM Variance bps", "body_black", numfmt_key="variance_bps")
    _xf("FBM Link",         "body_green", numfmt_key="number")
    _xf("FBM External",     "body_red", numfmt_key="number")
    _xf("FBM Assumption",   "body_blue_b", fill_key="yellow")
    _xf("FBM Total",        "body_black_b", border_key="total", numfmt_key="number")
    _xf("FBM Date",         "body_black", numfmt_key="date")
    _xf("FBM Year",         "body_black", numfmt_key="year", horiz="center")
    _xf("FBM Band",         "body_black", fill_key="band_grey")
    _xf("FBM KPI Big",      "kpi_big", horiz="center", vert="center")
    _xf("FBM KPI Label",    "units", horiz="center", vert="center")
    # Internal helpers used when building Config-style tabs by hand
    _xf("FBM Body Text C",  "body_black", horiz="center")
    _xf("FBM Body Text L",  "body_black", horiz="left")
    _xf("FBM Body Text Wrap", "body_black", horiz="left", vert="top", wrap=True)
    _xf("FBM Body $ Cents", "body_black", numfmt_key="currency2")

    xf_idx: Dict[str, int] = {name: cellxfs_count_before + i for i, (name, _) in enumerate(xf_specs)}
    close = _section_close(text, "cellXfs")
    text = text[:close] + "".join(xml for _, xml in xf_specs) + text[close:]
    text = _bump_count(text, "cellXfs", len(xf_specs))

    return text.encode("utf-8"), xf_idx


# ─── Sheet-body helpers (lxml-safe — Excel re-parses sheet bodies leniently) ─

def _ns():
    return "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def apply_cell_styles(
    sheet_bytes: bytes,
    coord_to_xfid: Mapping[str, int],
) -> bytes:
    """Update the ``s="N"`` attribute on specific cells of a worksheet.

    coord_to_xfid is a dict like ``{"A1": 155, "B3": 156, ...}`` where the
    values are ``cellXfs`` indexes (typically from ``extend_styles_bytewise``).

    Operates via lxml on the sheet body. This is safe because Excel re-parses
    sheet bodies leniently — unlike workbook.xml, styles.xml, and
    [Content_Types].xml, which are strict.
    """
    from lxml import etree
    NS = _ns()
    root = etree.fromstring(sheet_bytes)
    for c in root.findall(f".//{{{NS}}}c"):
        coord = c.get("r")
        if coord and coord in coord_to_xfid:
            c.set("s", str(coord_to_xfid[coord]))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def set_tab_color(sheet_bytes: bytes, hex_rgb: str) -> bytes:
    """Set or replace the tab color on a worksheet.

    ``hex_rgb`` is the 6-char ARGB color without the ``FF`` prefix
    (e.g. ``"1F4E79"`` for FBM input blue). Use the ``FBM_TAB`` constant
    dict for the four standard roles.
    """
    from lxml import etree
    NS = _ns()
    root = etree.fromstring(sheet_bytes)
    sp = root.find(f"{{{NS}}}sheetPr")
    if sp is None:
        sp = etree.Element(f"{{{NS}}}sheetPr")
        root.insert(0, sp)
    for tc in sp.findall(f"{{{NS}}}tabColor"):
        sp.remove(tc)
    etree.SubElement(sp, f"{{{NS}}}tabColor", rgb="FF" + hex_rgb)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def resolve_sheet_paths(workbook_xml: bytes, workbook_rels: bytes) -> Dict[str, str]:
    """Return a dict mapping sheet display name to its zip path.

    Excel may renumber sheet XML files on save (e.g. ``sheet15.xml`` becomes
    ``sheet13.xml``), so hardcoding paths from a previous build will break.
    Always resolve dynamically.
    """
    from lxml import etree
    NS = _ns()
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    wb = etree.fromstring(workbook_xml)
    rels = etree.fromstring(workbook_rels)
    rid_to_target = {r.get("Id"): r.get("Target") for r in rels}
    out: Dict[str, str] = {}
    for sh in wb.findall(f".//{{{NS}}}sheet"):
        name = sh.get("name")
        rid = sh.get(f"{{{REL_NS}}}id")
        if name and rid in rid_to_target:
            out[name] = "xl/" + rid_to_target[rid]
    return out


# ─── Self-test ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Sanity-check: extend a minimal styles.xml and verify lxml can re-parse
    # the result, with all 28 FBM styles correctly indexed.
    from lxml import etree

    minimal = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<fonts count="1"><font><sz val="11"/><color theme="1"/>'
        b'<name val="Calibri"/></font></fonts>'
        b'<fills count="2"><fill><patternFill patternType="none"/></fill>'
        b'<fill><patternFill patternType="gray125"/></fill></fills>'
        b'<borders count="1"><border/></borders>'
        b'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>'
        b'</styleSheet>'
    )
    new_bytes, xf = extend_styles_bytewise(minimal)
    # Verify byte-level envelope preserved
    assert new_bytes.startswith(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'), (
        "XML declaration must use double quotes and CRLF"
    )
    # Verify lxml parses it
    parsed = etree.fromstring(new_bytes)
    NS = _ns()
    cellxfs = parsed.find(f"{{{NS}}}cellXfs")
    # 1 original + 28 FBM = 29
    assert len(cellxfs) == 29, f"Expected 29 cellXfs, got {len(cellxfs)}"
    assert xf["FBM Title"] == 1, f"FBM Title should be index 1 (after the one original), got {xf['FBM Title']}"
    assert xf["FBM Header"] == 4
    print(f"surgical_styles self-test passed. {len(xf)} FBM styles injected.")
    print(f"FBM Title -> {xf['FBM Title']}, FBM Header -> {xf['FBM Header']}, FBM Formula $ -> {xf['FBM Formula $']}")
