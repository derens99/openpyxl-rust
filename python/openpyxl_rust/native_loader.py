"""Native formatting-preserving xlsx loader.

Parses the xlsx package directly (zipfile + ElementTree) so that
``load_workbook(path, data_only=False)`` preserves values, formulas, and
formatting without requiring openpyxl at runtime.
"""

import re
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Standard builtin number formats (ECMA-376 18.8.30)
_BUILTIN_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    41: r'_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
    42: r'_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
    43: r'_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',
    44: r'_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}

# Legacy indexed color palette (ECMA-376 18.8.27), alpha stripped.
_INDEXED_COLORS = [
    "000000",
    "FFFFFF",
    "FF0000",
    "00FF00",
    "0000FF",
    "FFFF00",
    "FF00FF",
    "00FFFF",
    "000000",
    "FFFFFF",
    "FF0000",
    "00FF00",
    "0000FF",
    "FFFF00",
    "FF00FF",
    "00FFFF",
    "800000",
    "008000",
    "000080",
    "808000",
    "800080",
    "008080",
    "C0C0C0",
    "808080",
    "9999FF",
    "993366",
    "FFFFCC",
    "CCFFFF",
    "660066",
    "FF8080",
    "0066CC",
    "CCCCFF",
    "000080",
    "FF00FF",
    "FFFF00",
    "00FFFF",
    "800080",
    "800000",
    "008080",
    "0000FF",
    "00CCFF",
    "CCFFFF",
    "CCFFCC",
    "FFFF99",
    "99CCFF",
    "FF99CC",
    "CC99FF",
    "FFCC99",
    "3366FF",
    "33CCCC",
    "99CC00",
    "FFCC00",
    "FF9900",
    "FF6600",
    "666699",
    "969696",
    "003366",
    "339966",
    "003300",
    "333300",
    "993300",
    "993366",
    "333399",
    "333333",
]

_DATE_FMT_STRIP_RE = re.compile(r'"[^"]*"|\[[^\]]*\]|\\.|AM/PM|am/pm')
_DATE_TOKEN_RE = re.compile(r"[dmhysDMHYS]")


def _is_date_format(fmt):
    if not fmt or fmt == "General":
        return False
    stripped = _DATE_FMT_STRIP_RE.sub("", fmt)
    # Reject text/number-only formats fast, then look for date tokens.
    return bool(_DATE_TOKEN_RE.search(stripped))


def _serial_to_pydate(serial):
    """Convert an Excel 1900-epoch serial number to datetime or time."""
    if 0 <= serial < 1:
        micros = round(serial * 86400 * 1_000_000)
        base = datetime(1970, 1, 1) + timedelta(microseconds=micros)
        return base.time()
    if serial >= 60:
        serial -= 1  # Excel's phantom 1900-02-29
    days = int(serial)
    frac = serial - days
    micros = round(frac * 86400 * 1_000_000)
    return datetime(1899, 12, 31) + timedelta(days=days, microseconds=micros)


def _color_from_el(el):
    """Convert a <color> element to a hex RGB string or None."""
    if el is None:
        return None
    rgb = el.get("rgb")
    if rgb:
        return rgb[2:] if len(rgb) == 8 else rgb
    indexed = el.get("indexed")
    if indexed is not None:
        idx = int(indexed)
        if 0 <= idx < len(_INDEXED_COLORS):
            return _INDEXED_COLORS[idx]
    return None  # theme/auto colors are not resolved


def _parse_fonts(styles_root):
    from openpyxl_rust.styles.fonts import Font

    fonts = []
    fonts_el = styles_root.find(f"{_MAIN}fonts")
    if fonts_el is None:
        return [None]
    for f in fonts_el.findall(f"{_MAIN}font"):
        name_el = f.find(f"{_MAIN}name")
        sz_el = f.find(f"{_MAIN}sz")
        u_el = f.find(f"{_MAIN}u")
        underline = None
        if u_el is not None:
            underline = u_el.get("val", "single")
            if underline == "none":
                underline = None
        vert_el = f.find(f"{_MAIN}vertAlign")
        fonts.append(
            Font(
                name=name_el.get("val") if name_el is not None else "Calibri",
                size=float(sz_el.get("val")) if sz_el is not None else 11,
                bold=f.find(f"{_MAIN}b") is not None,
                italic=f.find(f"{_MAIN}i") is not None,
                underline=underline,
                color=_color_from_el(f.find(f"{_MAIN}color")),
                strikethrough=f.find(f"{_MAIN}strike") is not None,
                vertAlign=vert_el.get("val") if vert_el is not None else None,
            )
        )
    return fonts


def _parse_fills(styles_root):
    from openpyxl_rust.styles.fills import PatternFill

    fills = []
    fills_el = styles_root.find(f"{_MAIN}fills")
    if fills_el is None:
        return [None]
    for f in fills_el.findall(f"{_MAIN}fill"):
        pattern = f.find(f"{_MAIN}patternFill")
        if pattern is None:
            fills.append(None)
            continue
        ptype = pattern.get("patternType")
        if ptype is None or ptype == "none":
            fills.append(None)
            continue
        fills.append(
            PatternFill(
                fill_type=ptype,
                start_color=_color_from_el(pattern.find(f"{_MAIN}fgColor")),
                end_color=_color_from_el(pattern.find(f"{_MAIN}bgColor")),
            )
        )
    return fills


def _parse_side(border_el, tag):
    from openpyxl_rust.styles.borders import Side

    el = border_el.find(f"{_MAIN}{tag}")
    if el is None:
        return Side()
    style = el.get("style")
    if style is None or style == "none":
        return Side()
    return Side(style=style, color=_color_from_el(el.find(f"{_MAIN}color")))


def _parse_borders(styles_root):
    from openpyxl_rust.styles.borders import Border

    borders = []
    borders_el = styles_root.find(f"{_MAIN}borders")
    if borders_el is None:
        return [None]
    for b in borders_el.findall(f"{_MAIN}border"):
        border = Border(
            left=_parse_side(b, "left"),
            right=_parse_side(b, "right"),
            top=_parse_side(b, "top"),
            bottom=_parse_side(b, "bottom"),
            diagonal=_parse_side(b, "diagonal"),
            diagonalUp=b.get("diagonalUp") in ("1", "true"),
            diagonalDown=b.get("diagonalDown") in ("1", "true"),
        )
        styled = any(s.style for s in (border.left, border.right, border.top, border.bottom, border.diagonal))
        borders.append(border if styled else None)
    return borders


def _parse_alignment(xf_el):
    from openpyxl_rust.styles.alignment import Alignment

    al = xf_el.find(f"{_MAIN}alignment")
    if al is None:
        return None
    return Alignment(
        horizontal=al.get("horizontal"),
        vertical=al.get("vertical"),
        wrap_text=al.get("wrapText") in ("1", "true"),
        shrink_to_fit=al.get("shrinkToFit") in ("1", "true"),
        indent=int(al.get("indent", 0)),
        text_rotation=int(al.get("textRotation", 0)),
    )


def _parse_styles(zf):
    """Parse xl/styles.xml into a list of per-xf style dicts."""
    number_formats = dict(_BUILTIN_FORMATS)
    xfs = []
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except KeyError:
        return xfs, number_formats

    numfmts_el = root.find(f"{_MAIN}numFmts")
    if numfmts_el is not None:
        for nf in numfmts_el.findall(f"{_MAIN}numFmt"):
            number_formats[int(nf.get("numFmtId"))] = nf.get("formatCode")

    fonts = _parse_fonts(root)
    fills = _parse_fills(root)
    borders = _parse_borders(root)

    cellxfs_el = root.find(f"{_MAIN}cellXfs")
    if cellxfs_el is not None:
        for xf in cellxfs_el.findall(f"{_MAIN}xf"):
            font_id = int(xf.get("fontId", 0))
            fill_id = int(xf.get("fillId", 0))
            border_id = int(xf.get("borderId", 0))
            numfmt_id = int(xf.get("numFmtId", 0))
            fmt = number_formats.get(numfmt_id, "General")
            xfs.append(
                {
                    "font": fonts[font_id] if 0 < font_id < len(fonts) else None,
                    "fill": fills[fill_id] if fill_id < len(fills) else None,
                    "border": borders[border_id] if border_id < len(borders) else None,
                    "alignment": _parse_alignment(xf),
                    "number_format": fmt if fmt != "General" else None,
                    "is_date": _is_date_format(fmt),
                }
            )
    return xfs, number_formats


def _parse_shared_strings(zf):
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings = []
    for si in root.findall(f"{_MAIN}si"):
        strings.append("".join(t.text or "" for t in si.iter(f"{_MAIN}t")))
    return strings


def _parse_rels(zf, rels_path):
    """Parse a .rels file into {rId: (target, target_mode, type)}."""
    rels = {}
    try:
        root = ET.fromstring(zf.read(rels_path))
    except KeyError:
        return rels
    for rel in root.findall(f"{_PKG_REL}Relationship"):
        rels[rel.get("Id")] = (
            rel.get("Target"),
            rel.get("TargetMode", "Internal"),
            rel.get("Type", ""),
        )
    return rels


def _resolve_target(target):
    """Resolve a workbook-relative rel target to a zip path."""
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    # Collapse any "xl/../" style segments
    parts = []
    for seg in target.split("/"):
        if seg == "..":
            if parts:
                parts.pop()
        else:
            parts.append(seg)
    return "/".join(parts)


def _cell_value(c_el, shared_strings, xf_is_date):
    """Extract a cell's Python value from its <c> element."""
    ctype = c_el.get("t", "n")
    f_el = c_el.find(f"{_MAIN}f")
    if f_el is not None:
        text = f_el.text or ""
        return "=" + text if text else None
    if ctype == "inlineStr":
        is_el = c_el.find(f"{_MAIN}is")
        if is_el is None:
            return None
        return "".join(t.text or "" for t in is_el.iter(f"{_MAIN}t"))
    v_el = c_el.find(f"{_MAIN}v")
    if v_el is None or v_el.text is None:
        return None
    raw = v_el.text
    if ctype == "s":
        return shared_strings[int(raw)]
    if ctype == "str" or ctype == "e":
        return raw
    if ctype == "b":
        return raw in ("1", "true")
    # numeric
    num = float(raw)
    if xf_is_date:
        return _serial_to_pydate(num)
    if num == int(num) and "e" not in raw.lower() and "." not in raw:
        return int(num)
    return num


def _parse_comments(zf, comments_path):
    """Parse a comments XML part into {ref: (text, author)}."""
    out = {}
    try:
        root = ET.fromstring(zf.read(comments_path))
    except KeyError:
        return out
    authors = [a.text or "" for a in root.iter(f"{_MAIN}author")]
    comment_list = root.find(f"{_MAIN}commentList")
    if comment_list is None:
        return out
    for comment in comment_list.findall(f"{_MAIN}comment"):
        ref = comment.get("ref")
        author_id = comment.get("authorId")
        author = authors[int(author_id)] if author_id is not None and authors else None
        text = "".join(t.text or "" for t in comment.iter(f"{_MAIN}t"))
        out[ref] = (text, author or None)
    return out


def _load_sheet(zf, ws, sheet_path, shared_strings, xfs):
    from openpyxl_rust.comments import Comment
    from openpyxl_rust.worksheet import _parse_cell_ref

    root = ET.fromstring(zf.read(sheet_path))

    # --- Tab color ---
    sheet_pr = root.find(f"{_MAIN}sheetPr")
    if sheet_pr is not None:
        tab_color = _color_from_el(sheet_pr.find(f"{_MAIN}tabColor"))
        if tab_color:
            ws.sheet_properties.tabColor = tab_color

    # --- Freeze panes ---
    for pane in root.iter(f"{_MAIN}pane"):
        if pane.get("state") in ("frozen", "frozenSplit"):
            top_left = pane.get("topLeftCell")
            if top_left:
                ws.freeze_panes = top_left

    # --- Column widths ---
    cols_el = root.find(f"{_MAIN}cols")
    if cols_el is not None:
        from openpyxl_rust.cell import _col_letter

        for col in cols_el.findall(f"{_MAIN}col"):
            width = col.get("width")
            if width is None:
                continue
            lo, hi = int(col.get("min")), int(col.get("max"))
            if hi - lo > 1000:
                hi = lo  # defensive cap against whole-sheet ranges
            for ci in range(lo, hi + 1):
                dim = ws.column_dimensions[_col_letter(ci)]
                dim.width = float(width)
                if col.get("hidden") in ("1", "true"):
                    dim.hidden = True

    # --- Cells ---
    sheet_data = root.find(f"{_MAIN}sheetData")
    if sheet_data is not None:
        for row_el in sheet_data.findall(f"{_MAIN}row"):
            ht = row_el.get("ht")
            if ht is not None:
                ws.row_dimensions[int(row_el.get("r"))].height = float(ht)
            if row_el.get("hidden") in ("1", "true"):
                ws.row_dimensions[int(row_el.get("r"))].hidden = True
            for c_el in row_el.findall(f"{_MAIN}c"):
                ref = c_el.get("r")
                if ref is None:
                    continue
                r, c = _parse_cell_ref(ref)
                style_idx = int(c_el.get("s", 0))
                xf = xfs[style_idx] if 0 < style_idx < len(xfs) else None
                value = _cell_value(c_el, shared_strings, xf["is_date"] if xf else False)
                if value is not None:
                    ws.cell(row=r, column=c, value=value)
                if xf and any((xf["font"], xf["fill"], xf["border"], xf["alignment"], xf["number_format"])):
                    cell = ws.cell(row=r, column=c)
                    if xf["font"] is not None:
                        cell.font = xf["font"]
                    if xf["fill"] is not None:
                        cell.fill = xf["fill"]
                    if xf["border"] is not None:
                        cell.border = xf["border"]
                    if xf["alignment"] is not None:
                        cell.alignment = xf["alignment"]
                    if xf["number_format"] is not None:
                        cell.number_format = xf["number_format"]

    # --- Merged cells ---
    merge_el = root.find(f"{_MAIN}mergeCells")
    if merge_el is not None:
        for mc in merge_el.findall(f"{_MAIN}mergeCell"):
            ws.merge_cells(mc.get("ref"))

    # --- Auto filter ---
    af = root.find(f"{_MAIN}autoFilter")
    if af is not None and af.get("ref"):
        ws.auto_filter.ref = af.get("ref")

    # --- Sheet protection ---
    prot = root.find(f"{_MAIN}sheetProtection")
    if prot is not None and prot.get("sheet") in ("1", "true"):
        ws.protection.sheet = True

    # --- Page setup ---
    ps = root.find(f"{_MAIN}pageSetup")
    if ps is not None:
        if ps.get("orientation"):
            ws.page_setup.orientation = ps.get("orientation")
        if ps.get("paperSize"):
            ws.page_setup.paperSize = int(ps.get("paperSize"))
        if ps.get("scale"):
            ws.page_setup.scale = int(ps.get("scale"))
        if ps.get("fitToWidth"):
            ws.page_setup.fitToWidth = int(ps.get("fitToWidth"))
        if ps.get("fitToHeight"):
            ws.page_setup.fitToHeight = int(ps.get("fitToHeight"))

    pm = root.find(f"{_MAIN}pageMargins")
    if pm is not None:
        for attr in ("left", "right", "top", "bottom", "header", "footer"):
            val = pm.get(attr)
            if val is not None:
                setattr(ws.page_margins, attr, float(val))

    # --- Hyperlinks and comments (need sheet rels) ---
    rels_path = sheet_path.rsplit("/", 1)[0] + "/_rels/" + sheet_path.rsplit("/", 1)[1] + ".rels"
    rels = _parse_rels(zf, rels_path)

    links_el = root.find(f"{_MAIN}hyperlinks")
    if links_el is not None:
        for link in links_el.findall(f"{_MAIN}hyperlink"):
            ref = link.get("ref")
            r, c = _parse_cell_ref(ref.split(":")[0])
            rid = link.get(f"{_REL}id")
            if rid and rid in rels:
                ws.cell(row=r, column=c).hyperlink = rels[rid][0]
            elif link.get("location"):
                ws.cell(row=r, column=c).hyperlink = "#" + link.get("location")

    for _rid, (target, _mode, rtype) in rels.items():
        if rtype.endswith("/comments"):
            comments = _parse_comments(zf, _resolve_target(target))
            for ref, (text, author) in comments.items():
                r, c = _parse_cell_ref(ref)
                ws.cell(row=r, column=c).comment = Comment(text, author)


def load_workbook_native(filename):
    """Load an xlsx with formatting into a Rust-backed Workbook."""
    from openpyxl_rust.workbook import DefinedName, Workbook
    from openpyxl_rust.worksheet import Worksheet

    zf = zipfile.ZipFile(filename)
    try:
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        wb_rels = _parse_rels(zf, "xl/_rels/workbook.xml.rels")
        shared_strings = _parse_shared_strings(zf)
        xfs, _number_formats = _parse_styles(zf)

        wb = Workbook()
        wb._sheets = []

        sheets_el = wb_root.find(f"{_MAIN}sheets")
        sheet_entries = []
        for sheet in sheets_el.findall(f"{_MAIN}sheet"):
            rid = sheet.get(f"{_REL}id")
            target = wb_rels.get(rid, (None,))[0]
            if target is None:
                continue
            path = _resolve_target(target)
            if "chartsheets" in path:
                continue  # chartsheets are not loadable worksheets
            sheet_entries.append((sheet.get("name"), path, sheet.get("state", "visible")))

        for i, (name, path, state) in enumerate(sheet_entries):
            if i == 0:
                sheet_idx = 0
                wb._rust_wb.set_sheet_title(0, name)
            else:
                sheet_idx = wb._rust_wb.add_sheet(name)
            ws = Worksheet(title=name, workbook=wb, sheet_idx=sheet_idx)
            if state in ("hidden", "veryHidden"):
                ws._sheet_state = state
            _load_sheet(zf, ws, path, shared_strings, xfs)
            wb._sheets.append(ws)

        # --- Defined names ---
        dn_el = wb_root.find(f"{_MAIN}definedNames")
        if dn_el is not None:
            for dn in dn_el.findall(f"{_MAIN}definedName"):
                if dn.get("name", "").startswith("_xlnm."):
                    continue  # builtin names (print areas etc.)
                wb.defined_names.add(DefinedName(dn.get("name"), attr_text=dn.text))

        return wb
    finally:
        zf.close()
