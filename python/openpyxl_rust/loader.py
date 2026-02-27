"""Formatting-preserving loader using real openpyxl for reading."""


def _convert_color(color_obj):
    """Convert an openpyxl Color object to a hex string or None."""
    if color_obj is None:
        return None
    # openpyxl Color has .rgb (ARGB string like "FF000000") or .theme or .indexed
    if hasattr(color_obj, "rgb") and color_obj.rgb and color_obj.rgb != "00000000":
        rgb = color_obj.rgb
        # openpyxl Color.rgb is sometimes a str like "FF0000FF" (ARGB)
        if isinstance(rgb, str) and len(rgb) == 8:
            return rgb[2:]  # strip alpha prefix
        elif isinstance(rgb, str) and len(rgb) == 6:
            return rgb
    return None


def _convert_font(src_font):
    """Convert an openpyxl Font to our Font."""
    from openpyxl_rust.styles.fonts import Font

    if src_font is None:
        return None

    color = None
    if src_font.color is not None:
        color = _convert_color(src_font.color)

    underline = None
    if src_font.underline is not None and src_font.underline != "none":
        underline = src_font.underline  # "single", "double", etc.

    return Font(
        name=src_font.name or "Calibri",
        size=src_font.size if src_font.size is not None else 11,
        bold=bool(src_font.bold),
        italic=bool(src_font.italic),
        underline=underline,
        color=color,
        strikethrough=bool(src_font.strikethrough),
        vertAlign=src_font.vertAlign,
    )


def _convert_alignment(src_align):
    """Convert an openpyxl Alignment to our Alignment."""
    from openpyxl_rust.styles.alignment import Alignment

    if src_align is None:
        return None

    return Alignment(
        horizontal=src_align.horizontal,
        vertical=src_align.vertical,
        wrap_text=bool(src_align.wrap_text) if src_align.wrap_text else False,
        shrink_to_fit=bool(src_align.shrink_to_fit) if src_align.shrink_to_fit else False,
        indent=int(src_align.indent) if src_align.indent else 0,
        text_rotation=int(src_align.text_rotation) if src_align.text_rotation else 0,
    )


def _convert_fill(src_fill):
    """Convert an openpyxl PatternFill to our PatternFill."""
    from openpyxl_rust.styles.fills import PatternFill

    if src_fill is None:
        return None

    # openpyxl has both PatternFill and GradientFill; we only support PatternFill
    fill_type = getattr(src_fill, "fill_type", None) or getattr(src_fill, "patternType", None)
    if fill_type is None or fill_type == "none":
        return None

    start_color = None
    end_color = None
    if hasattr(src_fill, "fgColor") and src_fill.fgColor:
        start_color = _convert_color(src_fill.fgColor)
    if hasattr(src_fill, "bgColor") and src_fill.bgColor:
        end_color = _convert_color(src_fill.bgColor)

    return PatternFill(
        fill_type=fill_type,
        start_color=start_color,
        end_color=end_color,
    )


def _convert_side(src_side):
    """Convert an openpyxl Side to our Side."""
    from openpyxl_rust.styles.borders import Side

    if src_side is None:
        return Side()

    style = src_side.style if src_side.style and src_side.style != "none" else None
    color = _convert_color(src_side.color) if src_side.color else None

    return Side(style=style, color=color)


def _convert_border(src_border):
    """Convert an openpyxl Border to our Border."""
    from openpyxl_rust.styles.borders import Border

    if src_border is None:
        return None

    return Border(
        left=_convert_side(src_border.left),
        right=_convert_side(src_border.right),
        top=_convert_side(src_border.top),
        bottom=_convert_side(src_border.bottom),
        diagonal=_convert_side(src_border.diagonal),
        diagonalUp=bool(src_border.diagonalUp) if src_border.diagonalUp else False,
        diagonalDown=bool(src_border.diagonalDown) if src_border.diagonalDown else False,
    )


def _has_formatting(src_cell):
    """Check if an openpyxl cell has any non-default formatting."""
    # Check font (non-default = bold, italic, color, size != 11, name != Calibri, etc.)
    f = src_cell.font
    if f and (
        f.bold
        or f.italic
        or f.strikethrough
        or f.underline
        or f.vertAlign
        or (f.color and _convert_color(f.color))
        or (f.name and f.name != "Calibri")
        or (f.size is not None and f.size != 11)
    ):
        return True

    # Check alignment
    a = src_cell.alignment
    if a and (
        a.horizontal
        or a.vertical
        or a.wrap_text
        or a.shrink_to_fit
        or (a.indent and a.indent != 0)
        or (a.text_rotation and a.text_rotation != 0)
    ):
        return True

    # Check fill
    fill = src_cell.fill
    if fill:
        ft = getattr(fill, "fill_type", None) or getattr(fill, "patternType", None)
        if ft and ft != "none":
            return True

    # Check border
    b = src_cell.border
    if b:
        for side in (b.left, b.right, b.top, b.bottom, b.diagonal):
            if side and side.style and side.style != "none":
                return True

    # Check number format
    nf = src_cell.number_format
    return bool(nf and nf != "General")


def _convert_openpyxl_to_rust(src_wb):
    """Convert a real openpyxl Workbook to our Rust-backed Workbook.

    Copies all cell values, formatting, merged cells, column widths,
    row heights, freeze panes, and other metadata.
    """
    from openpyxl_rust.comments import Comment
    from openpyxl_rust.workbook import Workbook
    from openpyxl_rust.worksheet import Worksheet

    wb = Workbook()
    wb._sheets = []

    for i, src_ws in enumerate(src_wb.worksheets):
        name = src_ws.title
        if i == 0:
            sheet_idx = 0
            wb._rust_wb.set_sheet_title(0, name)
        else:
            sheet_idx = wb._rust_wb.add_sheet(name)

        ws = Worksheet(title=name, workbook=wb, sheet_idx=sheet_idx)

        # --- Cell values and formatting ---
        for row in src_ws.iter_rows():
            for src_cell in row:
                if src_cell.value is None and not _has_formatting(src_cell):
                    continue

                r = src_cell.row
                c = src_cell.column

                # Set value
                val = src_cell.value
                if val is not None:
                    # Handle formula cells
                    if src_cell.data_type == "f" and isinstance(val, str) and val.startswith("="):
                        # Store as string for now (formulas not fully supported in our save)
                        ws.cell(row=r, column=c, value=val)
                    else:
                        ws.cell(row=r, column=c, value=val)
                else:
                    ws.cell(row=r, column=c)

                # Apply formatting only if cell has non-default formatting
                if _has_formatting(src_cell):
                    cell = ws.cell(row=r, column=c)

                    # Font
                    converted_font = _convert_font(src_cell.font)
                    if converted_font:
                        cell.font = converted_font

                    # Alignment
                    converted_align = _convert_alignment(src_cell.alignment)
                    if converted_align:
                        cell.alignment = converted_align

                    # Fill
                    converted_fill = _convert_fill(src_cell.fill)
                    if converted_fill:
                        cell.fill = converted_fill

                    # Border
                    converted_border = _convert_border(src_cell.border)
                    if converted_border:
                        cell.border = converted_border

                    # Number format
                    if src_cell.number_format and src_cell.number_format != "General":
                        cell.number_format = src_cell.number_format

                    # Hyperlink
                    if src_cell.hyperlink:
                        cell.hyperlink = src_cell.hyperlink.target

                    # Comment
                    if src_cell.comment:
                        cell.comment = Comment(
                            src_cell.comment.text,
                            src_cell.comment.author,
                        )

        # --- Column widths ---
        for col_letter, col_dim in src_ws.column_dimensions.items():
            if col_dim.width is not None and col_dim.width > 0:
                ws.column_dimensions[col_letter].width = col_dim.width

        # --- Row heights ---
        for row_num, row_dim in src_ws.row_dimensions.items():
            if row_dim.height is not None and row_dim.height > 0:
                ws.row_dimensions[row_num].height = row_dim.height

        # --- Freeze panes ---
        if src_ws.freeze_panes:
            ws.freeze_panes = src_ws.freeze_panes

        # --- Merged cells ---
        for merge_range in src_ws.merged_cells.ranges:
            ws.merge_cells(str(merge_range))

        # --- Print / page setup ---
        if src_ws.page_setup:
            ps = src_ws.page_setup
            if ps.orientation:
                ws.page_setup.orientation = ps.orientation
            if ps.paperSize is not None:
                ws.page_setup.paperSize = ps.paperSize
            if ps.scale is not None:
                ws.page_setup.scale = ps.scale
            if ps.fitToWidth is not None:
                ws.page_setup.fitToWidth = ps.fitToWidth
            if ps.fitToHeight is not None:
                ws.page_setup.fitToHeight = ps.fitToHeight

        if src_ws.page_margins:
            pm = src_ws.page_margins
            ws.page_margins.left = pm.left
            ws.page_margins.right = pm.right
            ws.page_margins.top = pm.top
            ws.page_margins.bottom = pm.bottom
            ws.page_margins.header = pm.header
            ws.page_margins.footer = pm.footer

        # --- Auto filter ---
        if src_ws.auto_filter and src_ws.auto_filter.ref:
            ws.auto_filter.ref = src_ws.auto_filter.ref

        # --- Protection ---
        if src_ws.protection and src_ws.protection.sheet:
            ws.protection.sheet = True
            ws.protection.password = getattr(src_ws.protection, "password", None)

        wb._sheets.append(ws)

    return wb
